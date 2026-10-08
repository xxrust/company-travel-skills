#!/usr/bin/env python
"""Prepare and optionally print invoice packets with chronological page labels."""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime
from pathlib import Path

import fitz


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


def chinese_font() -> str | None:
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    return None


def prepare_packet(inputs: list[Path], output: Path, *, kind: str, copies: int, trip_date: str) -> dict[str, object]:
    if not inputs:
        raise ValueError("至少提供一个 PDF 文件")
    if copies < 1:
        raise ValueError("份数必须大于等于 1")
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_suffix(".working.pdf")
    combined = fitz.open()
    source_map = []
    page_no = 0
    for src in inputs:
        doc = fitz.open(src)
        source_map.append({"source": str(src), "start_page": page_no + 1, "pages": len(doc)})
        combined.insert_pdf(doc)
        page_no += len(doc)
        doc.close()
    combined.save(temp)
    combined.close()

    stamped = output.with_suffix(".stamped.pdf")
    # Stamp the combined packet once, then duplicate it for the required copies.
    doc = fitz.open(temp)
    fontfile = chinese_font()
    for idx, page in enumerate(doc):
        label = f"{trip_date}  第 {idx + 1}/{len(doc)} 页"
        rect = fitz.Rect(page.rect.width - 190, page.rect.height - 28, page.rect.width - 12, page.rect.height - 8)
        page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
        kwargs = {"fontsize": 8, "align": fitz.TEXT_ALIGN_RIGHT, "color": (0, 0, 0), "overlay": True}
        if fontfile:
            kwargs["fontfile"] = fontfile
        page.insert_textbox(rect, label, **kwargs)
    doc.save(stamped)
    doc.close()
    temp.unlink(missing_ok=True)

    final_doc = fitz.open()
    for _ in range(copies):
        one = fitz.open(stamped)
        final_doc.insert_pdf(one)
        one.close()
    final_doc.save(output)
    pages = len(final_doc)
    final_doc.close()
    stamped.unlink(missing_ok=True)

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "trip_date": trip_date,
        "kind": kind,
        "copies": copies,
        "output": str(output),
        "pages_per_copy": pages // copies,
        "source_map": source_map,
    }
    output.with_suffix(".json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="生成带时间页码的发票打印包")
    parser.add_argument("inputs", nargs="*", help="输入 PDF，按粘贴顺序合并")
    parser.add_argument("--output", help="输出打印包 PDF")
    parser.add_argument("--kind", default="other", choices=["vat", "住宿", "hotel", "train", "火车", "flight", "机票", "车船费", "other"], help="票据类型")
    parser.add_argument("--copies", type=int, help="打印份数；省略时按票据类型自动决定")
    parser.add_argument("--trip-date", default=datetime.now().strftime("%Y-%m-%d"), help="右下角日期")
    parser.add_argument("--list-printers", action="store_true", help="列出本机打印机后退出")
    parser.add_argument("--printer", help="打印机名称；不提供则只生成文件")
    parser.add_argument("--print", action="store_true", dest="do_print", help="生成后调用 Windows 默认 PDF 打印关联程序打印")
    args = parser.parse_args()

    if args.list_printers:
        for item in discover_printers():
            print(f"{item['name']}\t默认={item['default']}")
        return 0

    if not args.inputs or not args.output:
        parser.error("生成打印包时必须提供输入 PDF 和 --output")

    copies = args.copies or default_copies(args.kind)
    manifest = prepare_packet([Path(x).expanduser().resolve() for x in args.inputs], Path(args.output).expanduser().resolve(), kind=args.kind, copies=copies, trip_date=args.trip_date)
    if args.printer:
        names = [x["name"] for x in discover_printers()]
        if args.printer not in names:
            raise ValueError(f"找不到打印机：{args.printer}；可用打印机：{names}")
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
