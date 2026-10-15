#!/usr/bin/env python
"""Pre-handover audit of a print packet or paper form.

Refuses to pass when content can be clipped by the printer, when the sheet is
lopsided, or when the footer label is missing. Run this before handing a packet
to the user for printing; a packet that fails must not be handed over.

Thresholds come from the target printer's measured device context: laser printers
typically leave a few millimetres unprintable on every side, so invoice content
must stay further inside than that (default 4.5mm, with an optional safety bump).

Usage:
    python scripts/check_packet.py <packet.pdf> [--min-mm 4.5] [--report out.txt]
"""

from __future__ import annotations

import argparse
from pathlib import Path

import fitz

MM = 72.0 / 25.4
REFUSE_BELOW_MM = 4.5


def audit(pdf: Path, min_mm: float = REFUSE_BELOW_MM, require_label: bool = True) -> tuple[bool, list[str]]:
    problems: list[str] = []
    lines: list[str] = []
    with fitz.open(pdf) as doc:
        if not len(doc):
            return False, [f"{pdf.name}: 空文件"]
        for index, page in enumerate(doc, start=1):
            items: list[tuple[float, float, float, float, bool]] = []
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        if text.strip():
                            x0, y0, x1, y1 = span["bbox"]
                            is_label = "第" in text and "/" in text
                            items.append((x0, y0, x1, y1, is_label))
            for drawing in page.get_drawings():
                rect = drawing.get("rect")
                if rect:
                    items.append((rect.x0, rect.y0, rect.x1, rect.y1, False))
            for image in page.get_images(full=True):
                for rect in page.get_image_rects(image[0]):
                    items.append((rect.x0, rect.y0, rect.x1, rect.y1, False))
            if not items:
                problems.append(f"sheet{index}: 空白页")
                continue

            content = [i for i in items if not i[4]]
            labels = [i for i in items if i[4]]
            if not content:
                problems.append(f"sheet{index}: 只有标注没有内容")
                continue
            left = min(i[0] for i in content) / MM
            right = (page.rect.width - max(i[2] for i in content)) / MM
            top = min(i[1] for i in content) / MM
            bottom = (page.rect.height - max(i[3] for i in items)) / MM
            lines.append(
                f"sheet{index}: 内容边距 左{left:.1f} 右{right:.1f} 上{top:.1f} 下{bottom:.1f} mm"
                f"  文本框{sum(1 for i in items if not i[4])} 标注{len(labels)}"
            )
            for name, value in (("左", left), ("右", right), ("上", top)):
                if value < min_mm:
                    problems.append(f"sheet{index}: {name}边距 {value:.1f}mm < 要求 {min_mm}mm（会被打印机裁掉）")
            if abs(left - right) > 4.0:
                problems.append(f"sheet{index}: 左右不对称（左{left:.1f} vs 右{right:.1f}mm）")
            if max(i[2] for i in items) >= page.rect.width - 0.01 or max(i[3] for i in items) >= page.rect.height - 0.01:
                problems.append(f"sheet{index}: 有元素贴到纸边（裁切特征）")
            if require_label and not labels:
                problems.append(f"sheet{index}: 缺少日期/页码标注")
    return (not problems), lines + problems


def main() -> int:
    parser = argparse.ArgumentParser(description="打印包交付前自审：内容是否会被裁、是否左右均衡、标注是否齐全")
    parser.add_argument("pdf")
    parser.add_argument("--min-mm", type=float, default=REFUSE_BELOW_MM, help=f"内容距纸边下限（默认 {REFUSE_BELOW_MM}mm）")
    parser.add_argument("--form", action="store_true", help="表单类文件（如报销单）：不要求日期/页码标注")
    parser.add_argument("--report", default="", help="结果写入该 UTF-8 文件")
    args = parser.parse_args()

    pdf = Path(args.pdf).expanduser().resolve()
    ok, lines = audit(pdf, args.min_mm, require_label=not args.form)
    header = f"file={pdf.name}  审核={'通过' if ok else '不通过'}"
    output = [header, *lines]
    text = "\n".join(output) + "\n"
    print(text, end="")
    if args.report:
        Path(args.report).expanduser().write_text(text, encoding="utf-8")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
