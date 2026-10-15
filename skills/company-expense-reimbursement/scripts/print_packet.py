#!/usr/bin/env python
"""Prepare and optionally print invoice packets with labels in blank areas.

Layout rules (company feedback):
  * two invoices share one A4 sheet (default ``--sheet 2up``),
  * the label goes into a reserved blank strip below the invoice, never on it,
  * content is anchored to the top-left of the sheet; it is never vertically
    centred, because that makes later cutting harder.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

import fitz


def expand_sources(args: list[str]) -> list[Path]:
    """Expand directories and glob patterns without relying on shell quoting.

    Round-tripping non-ASCII paths through some shells corrupts them, so the
    script expands a folder or a pattern itself.
    """
    found: list[Path] = []
    for raw in args:
        item = Path(raw).expanduser()
        if item.is_dir():
            for pattern in ("*.pdf", "*.ofd", "*.zip"):
                found += sorted(item.glob(pattern))
            continue
        if any(ch in raw for ch in "*?["):
            parent = item.parent if str(item.parent) else Path(".")
            found += sorted(parent.glob(item.name))
            continue
        found.append(item)
    seen: set[Path] = set()
    unique: list[Path] = []
    for path in found:
        resolved = path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(resolved)
    return unique


def pdf_from_ofd(ofd_path: Path, workdir: Path) -> Path | None:
    """Extract the PDF that every 12306 OFD archive carries alongside its XML."""
    try:
        with zipfile.ZipFile(ofd_path) as zf:
            candidates = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
            if not candidates:
                return None
            target = workdir / f"{ofd_path.stem}_{Path(candidates[0]).name}"
            target.write_bytes(zf.read(candidates[0]))
            return target
    except Exception:
        return None


def resolve_inputs(sources: list[Path], workdir: Path) -> tuple[list[tuple[Path, Path | None]], list[str]]:
    """Return [(printable PDF, source OFD if any)] plus notes on what was used."""
    items: list[tuple[Path, Path | None]] = []
    notes: list[str] = []
    for src in sources:
        suffix = src.suffix.lower()
        if suffix == ".pdf":
            items.append((src, None))
        elif suffix == ".ofd":
            extracted = pdf_from_ofd(src, workdir)
            if extracted:
                items.append((extracted, src))
                notes.append(f"{src.name}: 从 OFD 中提取同号 PDF 用于打印")
            else:
                notes.append(f"{src.name}: OFD 内未找到 PDF，已跳过（12306 的 zip 里通常带 PDF，优先传 zip）")
        elif suffix == ".zip":
            with zipfile.ZipFile(src) as zf:
                names = zf.namelist()
                # 12306 的 zip 同时带同号的 .ofd 与 .pdf；优先用 PDF 打印，
                # 但保留 OFD 以便读取结构化字段（如真实乘车日期）。
                pdf_names = [n for n in names if n.lower().endswith(".pdf")]
                ofd_names = [n for n in names if n.lower().endswith(".ofd")]
                written_pdf = written_ofd = None
                if pdf_names:
                    written_pdf = workdir / f"{src.stem}_{Path(pdf_names[0]).name}"
                    written_pdf.write_bytes(zf.read(pdf_names[0]))
                if ofd_names:
                    written_ofd = workdir / f"{src.stem}_{Path(ofd_names[0]).name}"
                    written_ofd.write_bytes(zf.read(ofd_names[0]))
                if written_pdf:
                    items.append((written_pdf, written_ofd))
                elif written_ofd:
                    extracted = pdf_from_ofd(written_ofd, workdir)
                    if extracted:
                        items.append((extracted, written_ofd))
                notes.append(f"{src.name}: 取用 {len(pdf_names) or len(ofd_names)} 个文件")
        else:
            notes.append(f"{src.name}: 不支持的类型，已跳过")
    return items, notes


def page_text(pdf: Path) -> str:
    try:
        with fitz.open(pdf) as doc:
            return "\n".join(page.get_text("text") for page in doc)
    except Exception:
        return ""


def abnormal_boxes(pdf: Path) -> bool:
    """True when a page's CropBox (or TrimBox) reaches outside its MediaBox.

    Some vendor e-invoices ship exactly that (MediaBox 595x842 with a CropBox
    610 wide starting at y=-22). Viewers take the intersection, which silently
    cuts off the left QR code and the right-hand invoice number before any
    printing happens.
    """
    try:
        with fitz.open(pdf) as doc:
            for page in doc:
                media, crop = page.mediabox, page.cropbox
                if (crop.x1 > media.x1 + 0.01 or crop.y1 > media.y1 + 0.01
                        or crop.x0 < media.x0 - 0.01 or crop.y0 < media.y0 - 0.01):
                    return True
    except Exception:
        return False
    return False


def normalize_boxes(pdf: Path, workdir: Path, pad: float | None = None,
                    min_paper_margin_mm: float = 5.0, packet_margin_pt: float = 10.0) -> tuple[Path, bool]:
    """Trim the visible area to the content plus a margin that survives printing.

    Two vendor problems are handled here, before any printing:

    * a CropBox reaching outside the MediaBox (viewers clip the intersection and
      silently drop the edges), and
    * content running flush to the page border (a hotel invoice measured only
      0.59mm of paper left on its right).

    The margin is derived, not fixed: a wide invoice is scaled down to fit the
    portrait A4 packet, so its source margin shrinks with it. ``pad`` is
    back-calculated so that after that scaling the content still keeps
    ``min_paper_margin_mm`` inside the paper edge. The MediaBox is widened when
    needed so the margin really exists. The source file is never modified.
    """
    with fitz.open(pdf) as doc:
        page = doc[0]
        media = page.mediabox
        xs, ys = [], []
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span.get("text", "").strip():
                        xs += [span["bbox"][0], span["bbox"][2]]
                        ys += [span["bbox"][1], span["bbox"][3]]
        for drawing in page.get_drawings():
            rect = drawing.get("rect")
            if rect:
                xs += [rect.x0, rect.x1]
                ys += [rect.y0, rect.y1]
        for image in page.get_images(full=True):
            for rect in page.get_image_rects(image[0]):
                xs += [rect.x0, rect.x1]
                ys += [rect.y0, rect.y1]
        if not xs:
            return pdf, False

        content_w = max(xs) - min(xs)
        content_h = max(ys) - min(ys)
        if pad is None:
            # 包内按 A4 双张排版：可用宽度 = A4 宽 - 左右留白 - 中间无（每张一行）
            slot_w = A4[0] - 2 * packet_margin_pt
            slot_h = (A4[1] - 2 * packet_margin_pt - 10.0) / 2 - 14.0
            scale = min(slot_w / max(content_w + 2 * 40.0, 1.0), slot_h / max(content_h + 2 * 40.0, 1.0))
            need_pt = min_paper_margin_mm / 25.4 * 72.0
            pad = max(6.0, need_pt / max(scale, 0.01))
        else:
            need_pt = pad
        pad = min(pad, 60.0)

        box = fitz.Rect(min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)
        # A page whose content already fills it needs no trimming at all.
        if (box.x0 <= 0.5 and box.y0 <= 0.5
                and box.x1 >= media.x1 - 0.5 and box.y1 >= media.y1 - 0.5):
            return pdf, False
        # Raw PDF boxes use a bottom-left origin; PyMuPDF reports top-left page
        # coordinates, so convert before writing. MediaBox must enclose CropBox.
        media_box = (
            f"[{min(media.x0, box.x0) - 1:.2f} {min(media.y0, media.y1 - box.y1) - 1:.2f} "
            f"{max(media.x1, box.x1) + 1:.2f} {max(media.y0, media.y1 - box.y0) + 1:.2f}]"
        )
        crop_box = f"[{box.x0:.2f} {media.y1 - box.y1:.2f} {box.x1:.2f} {media.y1 - box.y0:.2f}]"
    data = pdf.read_bytes()
    data = re.sub(rb"/MediaBox\s*\[[^\]]*\]", b"/MediaBox " + media_box.encode(), data, count=1)
    for key in (b"/CropBox", b"/TrimBox"):
        if key in data:
            data = re.sub(key + rb"\s*\[[^\]]*\]", key + b" " + crop_box.encode(), data, count=1)
            break
    target = workdir / f"normalized_{pdf.name}"
    target.write_bytes(data)
    return target, True


def is_supplement_doc(pdf: Path) -> bool:
    """True for a supporting document (行程单/结账单) rather than the invoice itself.

    An invoice always carries an invoice number; a trip table or hotel settlement
    statement never does, even though it may echo a 价税合计 line.
    """
    return "发票号码" not in page_text(pdf)


def route_keys(pdf: Path) -> set[str]:
    """Evidence keys this document carries: travel date, endpoints, stay dates."""
    text = page_text(pdf)
    keys: set[str] = set()
    if any(k in text for k in ("出行日期", "乘车日期")):
        keys.add("travel_date")
    if any(k in text for k in ("入离日期", "入住日期", "离店日期")):
        keys.add("stay_dates")
    if ("出发地" in text and "到达地" in text) or ("出发站" in text and "到达站" in text):
        keys.add("endpoints")
    return keys


def invoice_has_route(pdf: Path) -> bool:
    """True when the invoice itself states the travel date and both endpoints."""
    keys = route_keys(pdf)
    return "travel_date" in keys and "endpoints" in keys


def prune_supplements(items: list[tuple[Path, Path | None]]) -> tuple[list[tuple[Path, Path | None]], list[str]]:
    """Drop supporting documents whose facts the invoices already state.

    Trip tables are kept only when an invoice lacks the travel date or the two
    endpoints; hotel settlement statements are always supporting documents and are
    never printed. Missing route evidence never blocks the packet.
    """
    invoices = [item for item in items if not is_supplement_doc(item[0])]
    supplements = [item for item in items if is_supplement_doc(item[0])]
    if not invoices:
        # No recognisable invoice (e.g. bare PDFs for testing): print what was given.
        return items, []

    notes = []
    keep: list[tuple[Path, Path | None]] = []
    for pdf, ofd in supplements:
        needs_route = any("travel_date" not in route_keys(inv) or "endpoints" not in route_keys(inv) for inv, _ in invoices)
        text = page_text(pdf)
        is_trip_table = "行程单" in text or "行程起止日期" in text
        if is_trip_table and needs_route:
            keep.append((pdf, ofd))
            notes.append(f"{pdf.name}: 发票缺少行程说明，保留行程单作为补充")
        else:
            notes.append(f"{pdf.name}: 辅助凭证（{'行程单' if is_trip_table else '结账单'}），不打印、不重复计入")
    return invoices + keep, notes


def travel_dates_ofd(ofd_path: Path | None) -> list[str]:
    """Read 乘车日期 (TravelDate) from the XBRL invoice XML inside a 12306 OFD."""
    if not ofd_path or not ofd_path.exists():
        return []
    try:
        from xml.etree import ElementTree as ET

        with zipfile.ZipFile(ofd_path) as zf:
            entry = [n for n in zf.namelist() if n.startswith("Doc_0/Attachs/rai_") and n.endswith(".xml")]
            if not entry:
                return []
            root = ET.fromstring(zf.read(entry[0]))
        return sorted({(node.text or "").strip() for node in root.iter() if node.tag.endswith("}TravelDate") and (node.text or "").strip()})
    except Exception:
        return []


def travel_dates_pdf(pdf: Path) -> list[str]:
    """Dates printed on the ticket, preferring explicit travel-date labels.

    A generic date sweep also picks up the issue date (开票日期), which must not
    end up on the label, so labelled fields win and every 开票日期 is excluded.
    """
    text = page_text(pdf)

    def collect(segment: str) -> set[str]:
        days = {f"{int(y):04d}-{int(m):02d}-{int(d):02d}" for y, m, d in re.findall(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", segment)}
        days |= {f"{int(y):04d}-{int(m):02d}-{int(d):02d}" for y, m, d in re.findall(r"(20\d{2})-(\d{1,2})-(\d{1,2})", segment)}
        return days

    issue_dates = set()
    for match in re.finditer("开票日期", text):
        issue_dates |= collect(text[match.start():match.start() + 40])

    labelled: set[str] = set()
    # 交通票用出行/乘车日期；住宿票用入离日期（或入住/离店日期）。
    for key in ("出行日期", "乘车日期", "行程起止日期", "乘车时间", "入离日期", "入住日期", "离店日期"):
        for match in re.finditer(re.escape(key), text):
            labelled |= collect(text[match.start():match.start() + 120])
    # 入离日期:10-12至10-13 这类不带年份的写法，补上开票年份。
    for match in re.finditer(r"入离日期[:：]?\s*(\d{1,2})-(\d{1,2})\s*至\s*(\d{1,2})-(\d{1,2})", text):
        year_match = re.search(r"开票日期[:：]?\s*(20\d{2})", text)
        year = int(year_match.group(1)) if year_match else datetime.now().year
        for month, day in ((match.group(1), match.group(2)), (match.group(3), match.group(4))):
            labelled.add(f"{year:04d}-{int(month):02d}-{int(day):02d}")
    labelled -= issue_dates
    if labelled:
        return sorted(labelled)

    days = collect(text) - issue_dates
    return sorted(days)


def detect_trip_date(items: list[tuple[Path, Path | None]]) -> str | None:
    """Label date: the travel-date range across the packet's invoices."""
    days: set[str] = set()
    for pdf, ofd in items:
        found = travel_dates_ofd(ofd) or travel_dates_pdf(pdf)
        days |= set(found)
    if not days:
        return None
    ordered = sorted(days)
    return ordered[0] if len(ordered) == 1 else f"{ordered[0]} 至 {ordered[-1]}"


def default_copies(kind: str) -> int:
    kind = kind.lower()
    # 结账单只作辅助凭证，不打印
    if kind in {"结账单", "settlement", "checkout"}:
        return 0
    if kind in {"vat", "住宿", "hotel", "train", "火车", "flight", "机票", "车船费"}:
        return 2
    return 1


def discover_printers() -> list[dict[str, object]]:
    try:
        import win32print
    except ImportError as exc:
        raise RuntimeError("缺少 pywin32，无法读取 Windows 打印机") from exc
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    rows = win32print.EnumPrinters(flags)
    default_name = win32print.GetDefaultPrinter()
    printers = []
    for row in rows:
        name = row[2]
        printers.append({"name": name, "default": name == default_name})
    return printers


def chinese_font() -> tuple[str, str] | None:
    """Return (fontname, fontfile) pairs that embed CJK glyphs reliably.

    A .ttc collection must not be registered through the built-in CJK
    "china-s" font: that combination renders 中文 as "?" or with wide gaps.
    Pairing a TTF/TTC file with its own font name embeds the real glyphs.
    """
    candidates = [
        ("simhei", r"C:\Windows\Fonts\simhei.ttf"),
        ("simsun", r"C:\Windows\Fonts\simsun.ttc"),
        ("msyh", r"C:\Windows\Fonts\msyh.ttc"),
    ]
    for name, path in candidates:
        if Path(path).exists():
            return name, path
    return None


A4 = (595.276, 841.890)


def _slot_geometry(per_sheet: int, margin: float, gap: float, label_strip: float) -> tuple[float, float, float, float]:
    """Return (origin_x, content_w, content_h, slot_pitch) for a symmetric A4 layout."""
    safe_w = A4[0] - 2 * margin
    safe_h = A4[1] - 2 * margin
    slot_h = (safe_h - gap * (per_sheet - 1)) / per_sheet
    return margin, safe_w, slot_h - label_strip, slot_h + gap


def content_bbox(page: fitz.Page, pad: float = 3.0) -> fitz.Rect:
    """Bounding box of what the invoice actually prints.

    Invoices (rail e-tickets in particular) carry a wide blank area inside their
    own page. Drawing that blank area makes the paper look lopsided even when the
    frames are symmetric, so the packet is placed by content, not by page box.
    """
    xs: list[float] = []
    ys: list[float] = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if span.get("text", "").strip():
                    x0, y0, x1, y1 = span["bbox"]
                    xs += [x0, x1]
                    ys += [y0, y1]
    for image in page.get_images(full=True):
        for rect in page.get_image_rects(image[0]):
            xs += [rect.x0, rect.x1]
            ys += [rect.y0, rect.y1]
    for drawing in page.get_drawings():
        rect = drawing.get("rect")
        if rect:
            xs += [rect.x0, rect.x1]
            ys += [rect.y0, rect.y1]
    if not xs:
        return page.rect
    box = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
    box = fitz.Rect(box.x0 - pad, box.y0 - pad, box.x1 + pad, box.y1 + pad)
    return box & page.rect


def _draw_page(sheet: fitz.Page, page: fitz.Page, target: fitz.Rect, crop: fitz.Rect,
               rotate: int = 0) -> None:
    """Place a (optionally cropped, optionally rotated) source page into ``target``."""
    if rotate:
        base = page.rotation_matrix * fitz.Matrix(rotate)
        if rotate in (90, 270):
            base = base * fitz.Matrix(1, -1)
        src = crop * base
        matrix = base * fitz.Matrix(1, 0, 0, 1, -src.x0, -src.y0) * fitz.Matrix(
            target.width / src.width, 0, 0, target.height / src.height, target.x0, target.y0
        )
        sheet.show_pdf_page(target, page.parent, page.number, clip=src, rotate=rotate)
        return
    sheet.show_pdf_page(target, page.parent, page.number, clip=crop)
    # 不要在这里补白色遮罩：show_pdf_page 的 clip 已经裁掉了范围外的内容，
    # 再画白条会盖住发票自己的内容（曾把滴滴的二维码、住宿发票左侧竖排的
    # "购买方信息"标签盖掉）。


def build_2up(pages: list[fitz.Page], *, margin: float = 10.0, gap: float = 10.0,
              label_strip: float = 14.0, trip_date: str, scale_to_fit: bool = True,
              cut_guides: bool = False, crop: bool = True) -> fitz.Document:
    """Place two invoice pages per A4 sheet, one above the other, top-anchored.

    Deliberately plain: symmetric margins, content anchored at the top-left of
    each slot, label in a reserved blank strip below the invoice. The printer
    driver's own "fit to page" handles the final scaling, so no printer-specific
    offsets are baked in here.
    """
    out = fitz.open()
    font = chinese_font()
    total = len(pages)
    per_sheet = 2
    origin_x, content_w, content_h, pitch = _slot_geometry(per_sheet, margin, gap, label_strip)
    origin_y = margin
    width, height = A4

    for index in range(0, total, per_sheet):
        sheet = out.new_page(width=width, height=height)
        for slot, page in enumerate(pages[index:index + per_sheet]):
            slot_top = origin_y + slot * pitch
            crop_rect = content_bbox(page) if crop else page.rect
            src_w, src_h = crop_rect.width, crop_rect.height
            scale = min(content_w / src_w, content_h / src_h)
            if not scale_to_fit:
                scale = 1.0
            draw_w = src_w * scale
            draw_h = src_h * scale
            target = fitz.Rect(origin_x, slot_top, origin_x + draw_w, slot_top + draw_h)
            _draw_page(sheet, page, target, crop_rect)
            if cut_guides:
                sheet.draw_rect(target, color=(0.6, 0.6, 0.6), width=0.4, overlay=True)
            label = f"{trip_date}  第 {index + slot + 1}/{total} 页"
            label_rect = fitz.Rect(
                origin_x,
                slot_top + draw_h + 1,
                origin_x + content_w,
                slot_top + draw_h + label_strip,
            )
            kwargs = {"fontsize": 8, "align": fitz.TEXT_ALIGN_RIGHT, "color": (0, 0, 0), "overlay": True}
            if font:
                kwargs["fontname"], kwargs["fontfile"] = font
            sheet.insert_textbox(label_rect, label, **kwargs)
    return out


def build_invoice_packet(pages: list[fitz.Page], *, trip_date: str, per_sheet: int = 2,
                         margin: float = 10.0, gap: float = 10.0, label_strip: float = 14.0,
                         min_scale: float = 0.55, cut_guides: bool = False,
                         rotate_wide: bool = False, copies: int = 1,
                         number_base: int = 1, number_total: int | None = None,
                         slot_labels: list[int] | None = None) -> tuple[fitz.Document, int]:
    """Lay invoice pages onto A4 sheets, cropping each to its content.

    ``rotate_wide`` is off by default: rotating a wide invoice sideways made the
    sheet harder to read and cut the second slot, so landscape invoices stay
    upright and are simply scaled to fit the portrait slot.
    Defaults to ``per_sheet`` invoices per sheet; when that would shrink a page
    below ``min_scale`` (small fonts stop being readable), the packet falls back
    to one invoice per sheet. Returns (document, invoices per sheet).

    ``number_base`` / ``number_total`` place this packet's first invoice inside the
    case-wide paste order, so a per-kind packet still prints the same 第 x/N 页 as
    the merged packet.
    """
    if not pages:
        raise ValueError("没有可排版的发票页")

    def rotation_of(page: fitz.Page) -> int:
        if not rotate_wide:
            return 0
        return 90 if page.rect.width > page.rect.height else 0

    rotations = [rotation_of(page) for page in pages]
    crops = []
    for page, angle in zip(pages, rotations):
        box = content_bbox(page)
        if angle:
            base = page.rotation_matrix * fitz.Matrix(angle)
            if angle in (90, 270):
                base = base * fitz.Matrix(1, -1)
            box = box * base
        crops.append(box)

    def fits(slots: int) -> bool:
        content_w, content_h = _slot_geometry(slots, margin, gap, label_strip)[1:3]
        for crop in crops:
            if min(content_w / crop.width, content_h / crop.height) < min_scale:
                return False
        return True

    slots = max(1, per_sheet)
    while slots > 1 and not fits(slots):
        slots -= 1

    out = fitz.open()
    font = chinese_font()
    total = len(pages)
    origin_x, content_w, content_h, pitch = _slot_geometry(slots, margin, gap, label_strip)
    origin_y = margin
    width, height = A4

    for start in range(0, total, slots):
        sheet = out.new_page(width=width, height=height)
        for slot, page in enumerate(pages[start:start + slots]):
            crop = crops[start + slot]
            slot_top = origin_y + slot * pitch
            scale = min(content_w / crop.width, content_h / crop.height)
            draw_w = crop.width * scale
            draw_h = crop.height * scale
            target = fitz.Rect(origin_x, slot_top, origin_x + draw_w, slot_top + draw_h)
            _draw_page(sheet, page, target, crop, rotate=rotations[start + slot])
            if cut_guides:
                sheet.draw_rect(target, color=(0.6, 0.6, 0.6), width=0.4, overlay=True)
            # 页码按"发票张数"递增：同一张发票的多份共用同一序号，否则会印出
            # 第7/6页 这类越界编号。``slot_labels`` 允许由编排脚本按时间线直接
            # 指定每张发票的编号（分包内部顺序与整案顺序不一致时必须用它）。
            invoices_per_copy = max(1, total // max(1, copies))
            invoice_index = (start + slot) // max(1, copies)
            shown_total = number_total or invoices_per_copy
            if slot_labels and invoice_index < len(slot_labels):
                number = slot_labels[invoice_index]
            else:
                number = number_base + invoice_index
            label = f"{trip_date}  第 {number}/{shown_total} 页"
            label_rect = fitz.Rect(
                origin_x,
                slot_top + draw_h + 1,
                origin_x + content_w,
                slot_top + draw_h + label_strip,
            )
            kwargs = {"fontsize": 8, "align": fitz.TEXT_ALIGN_RIGHT, "color": (0, 0, 0), "overlay": True}
            if font:
                kwargs["fontname"], kwargs["fontfile"] = font
            sheet.insert_textbox(label_rect, label, **kwargs)
    return out, slots


def build_copies_packet(specs: list[tuple[Path, str]], *, margin: float = 28.0, gap: float = 10.0,
                        label_strip: float = 14.0, per_sheet: int = 2,
                        min_scale: float = 0.55) -> fitz.Document:
    """Place labelled invoice copies onto A4 sheets, ``per_sheet`` per sheet.

    Each entry is ``(source PDF, label text)`` and represents ONE physical copy —
    the reimbursement copy (报销联) or the deduction copy (抵扣联) of an invoice.
    The two sequences are numbered independently, so the caller decides the label
    text; this function only lays them out top-anchored with symmetric margins and
    writes the label into the reserved blank strip below each invoice.

    The label is pulled well inside the right edge: a right-aligned label ending at
    the sheet margin landed ~5.6mm from the paper border and was clipped by the
    printer's unprintable edge, which looked like the invoice itself was cut.
    """
    if not specs:
        raise ValueError("没有可排版的发票副本")
    font = chinese_font()
    origin_x, content_w, content_h, pitch = _slot_geometry(per_sheet, margin, gap, label_strip)
    origin_y = margin
    width, height = A4
    out = fitz.open()
    label_right_inset = 40.0   # 标注右端内缩 40pt(14mm)，确保落纸后远离不可打印边

    for start in range(0, len(specs), per_sheet):
        sheet = out.new_page(width=width, height=height)
        for slot, (pdf, label) in enumerate(specs[start:start + per_sheet]):
            with fitz.open(pdf) as doc:
                src_page = doc[0]
                crop = content_bbox(src_page)
                scale = min(content_w / crop.width, content_h / crop.height)
                draw_w = crop.width * scale
                draw_h = crop.height * scale
                slot_top = origin_y + slot * pitch
                target = fitz.Rect(origin_x, slot_top, origin_x + draw_w, slot_top + draw_h)
                sheet.show_pdf_page(target, src_page.parent, src_page.number, clip=crop)
            label_rect = fitz.Rect(origin_x, slot_top + draw_h + 1,
                                   origin_x + content_w - label_right_inset,
                                   slot_top + draw_h + label_strip)
            kwargs = {"fontsize": 8, "align": fitz.TEXT_ALIGN_RIGHT, "color": (0, 0, 0), "overlay": True}
            if font:
                kwargs["fontname"], kwargs["fontfile"] = font
            sheet.insert_textbox(label_rect, label, **kwargs)
    return out


def build_1up(pages: list[fitz.Page], *, trip_date: str) -> fitz.Document:
    """Keep one invoice per page, with the label in a reserved strip below it."""
    out = fitz.open()
    font = chinese_font()
    total = len(pages)
    for index, page in enumerate(pages):
        width, height = page.rect.width, page.rect.height
        strip = 14.0
        sheet = out.new_page(width=width, height=height + strip)
        sheet.show_pdf_page(fitz.Rect(0, 0, width, height), page.parent, page.number)
        label = f"{trip_date}  第 {index + 1}/{total} 页"
        label_rect = fitz.Rect(0, height + 1, width - 6, height + strip)
        kwargs = {"fontsize": 8, "align": fitz.TEXT_ALIGN_RIGHT, "color": (0, 0, 0), "overlay": True}
        if font:
            kwargs["fontname"], kwargs["fontfile"] = font
        sheet.insert_textbox(label_rect, label, **kwargs)
    return out


def prepare_packet(inputs: list[Path], output: Path, *, kind: str, copies: int, trip_date: str,
                   sheet: str = "2up", margin: float = 10.0, scale_to_fit: bool = True,
                   cut_guides: bool = False, crop: bool = True) -> dict[str, object]:
    if not inputs:
        raise ValueError("至少提供一个 PDF 文件")
    if copies < 1:
        raise ValueError("份数必须大于等于 1")
    output.parent.mkdir(parents=True, exist_ok=True)

    source = fitz.open()
    source_map = []
    start = 0
    for src in inputs:
        doc = fitz.open(src)
        source_map.append({"source": str(src), "start_page": start + 1, "pages": len(doc)})
        source.insert_pdf(doc)
        start += len(doc)
        doc.close()

    # Copies: every copy repeats the full invoice sequence.
    pages = [source[i] for i in range(len(source))] * copies
    if sheet == "2up":
        final = build_2up(pages, margin=margin, trip_date=trip_date, scale_to_fit=scale_to_fit,
                          cut_guides=cut_guides, crop=crop)
    else:
        final = build_1up(pages, trip_date=trip_date)
    final.save(output)
    sheets = len(final)
    final.close()
    source.close()

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "trip_date": trip_date,
        "kind": kind,
        "copies": copies,
        "sheet": sheet,
        "invoices": len(inputs) * copies,
        "sheets": sheets,
        "margin_pt": margin,
        "cut_guides": cut_guides,
        "output": str(output),
        "source_map": source_map,
    }
    output.with_suffix(".json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="生成带日期页码的发票打印包（默认一张A4放两张，标注落在空白处）")
    parser.add_argument("inputs", nargs="*", help="输入 PDF，按粘贴顺序合并")
    parser.add_argument("--output", help="输出打印包 PDF")
    parser.add_argument("--kind", default="other", choices=["vat", "住宿", "hotel", "train", "火车", "flight", "机票", "车船费", "other"], help="票据类型")
    parser.add_argument("--copies", type=int, help="打印份数；省略时按票据类型自动决定")
    parser.add_argument("--sheet", choices=["2up", "1up"], default="2up", help="2up=一张A4两张发票（默认）；1up=一张A4一张发票")
    parser.add_argument("--margin", type=float, default=10.0, help="2up 页面边距（点）")
    parser.add_argument("--no-scale", action="store_true", dest="no_scale", help="2up 时不缩放原尺寸（可能超出边距）")
    parser.add_argument("--cut-guides", action="store_true", dest="cut_guides", help="在每张发票外框绘制细灰裁剪参考线")
    parser.add_argument("--no-crop", action="store_true", dest="no_crop", help="不按内容边界裁剪发票自带空白（默认会裁剪，使左右留白对称）")
    parser.add_argument("--trip-date", default=datetime.now().strftime("%Y-%m-%d"), help="右下角日期")
    parser.add_argument("--list-printers", action="store_true", help="列出本机打印机后退出")
    parser.add_argument("--printer", help="打印机名称；提供后校验是否存在（排版不依赖打印机）")
    parser.add_argument("--print", action="store_true", dest="do_print", help="生成后交给系统 PDF 打印关联程序")
    args = parser.parse_args()

    if args.list_printers:
        for item in discover_printers():
            print(f"{item['name']}\t默认={item['default']}")
        return 0

    if not args.inputs or not args.output:
        parser.error("生成打印包时必须提供输入 PDF 和 --output")

    copies = args.copies or default_copies(args.kind)
    if args.printer:
        names = [x["name"] for x in discover_printers()]
        if args.printer not in names:
            raise ValueError(f"找不到打印机：{args.printer}；可用打印机：{names}")
    manifest = prepare_packet(
        [Path(x).expanduser().resolve() for x in args.inputs],
        Path(args.output).expanduser().resolve(),
        kind=args.kind,
        copies=copies,
        trip_date=args.trip_date,
        sheet=args.sheet,
        margin=args.margin,
        scale_to_fit=not args.no_scale,
        cut_guides=args.cut_guides,
        crop=not args.no_crop,
    )
    if args.do_print:
        if args.printer:
            # The Windows shell printto verb passes the selected printer name to
            # the installed PDF handler without changing the system default.
            os.startfile(str(Path(args.output).resolve()), "printto", f'"{args.printer}"')
        else:
            os.startfile(str(Path(args.output).resolve()), "print")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
