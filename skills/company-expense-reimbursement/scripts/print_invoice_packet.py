#!/usr/bin/env python
"""Invoice print packet (通用发票打印包).

Frozen layout, confirmed on the real print workflow:

  * two same-size invoices per A4 sheet, one above the other; when that would
    shrink a page below readability (or the page is A4 portrait, e.g. a trip
    table), it automatically falls back to one invoice per sheet,
  * every invoice is cropped to its printed content, so the paper keeps a
    symmetric margin around it,
  * the date/page label sits in a reserved blank strip below the invoice and
    never overlaps it,
  * content is anchored top-left; nothing is ever auto-centred,
  * no printer-specific offsets or scaling: the printer's own "fit to page"
    performs the final scaling, so changing printers needs no new parameters,
  * 行程单/结账单 are dropped when the invoices already state the travel date and
    both endpoints, and kept when they do not.

Usage:
    python scripts/print_invoice_packet.py <文件|目录|通配符> ... \
        --output "05-报销/发票打印包.pdf" [--kind taxi] [--copies 2] [--trip-date 2026-10-13]
"""

from __future__ import annotations

import argparse
import tempfile
from datetime import datetime
from pathlib import Path

import fitz

from print_packet import (
    build_invoice_packet,
    default_copies,
    detect_trip_date,
    expand_sources,
    normalize_boxes,
    prune_supplements,
    resolve_inputs,
)

DEFAULT_MARGIN = 17.0  # 6.0mm：必须大于打印机 4.23mm 不可打印边，10pt(3.5mm) 会被切
DEFAULT_GAP = 10.0


def main() -> int:
    parser = argparse.ArgumentParser(description="生成发票打印包（一张A4两张、对称留白、标注在空白处、不居中）")
    parser.add_argument("inputs", nargs="+", help="发票文件、目录或通配符（PDF / OFD / 12306 的 zip）")
    parser.add_argument("--output", help="输出打印包 PDF（可与 --name 二选一）")
    parser.add_argument("--output-dir", default="", help="输出目录；配合 --name 使用，推荐（避免非 ASCII 路径经 shell 传参被破坏）")
    parser.add_argument("--name", default="", help="输出文件名，如 invoice-packet-taxi.pdf")
    parser.add_argument("--kind", default="other",
                        help="票种：train/taxi/flight/hotel/vat…（决定份数：住宿与交通 2 份，其他 1 份）")
    parser.add_argument("--copies", type=int, help="每张发票份数；省略时按 --kind 自动决定")
    parser.add_argument("--trip-date", default="", help="标注日期；省略时自动取发票上的乘车/出行日期")
    parser.add_argument("--per-sheet", type=int, default=2, help="每张 A4 目标发票数（默认 2；放不下会自动降为 1）")
    parser.add_argument("--margin", type=float, default=DEFAULT_MARGIN, help=f"四周留白（点，默认 {DEFAULT_MARGIN}）")
    parser.add_argument("--cut-guides", action="store_true", dest="cut_guides", help="发票外框加细灰裁剪参考线")
    parser.add_argument("--report", default="", help="把结果写入该 UTF-8 文件（规避控制台非 UTF-8 编码）")
    parser.add_argument("--start-number", type=int, default=1,
                        help="本包第一张发票在整案粘贴顺序中的序号（默认 1；分包共用一套编号时指定）")
    parser.add_argument("--total-number", type=int, default=0,
                        help="整案发票张数（默认本包张数）；分包时指定，使页码形如 第3/6页")
    parser.add_argument("--slot-labels", default="",
                        help="按本包内发票顺序指定编号，逗号分隔，如 2,3,6（分包顺序与整案时间线不一致时使用）")
    args = parser.parse_args()

    if args.output:
        output = Path(args.output).expanduser().resolve()
    elif args.output_dir and args.name:
        output = (Path(args.output_dir).expanduser() / args.name).resolve()
    else:
        parser.error("需要 --output，或同时提供 --output-dir 与 --name")
    if output.exists():
        raise SystemExit(f"拒绝覆盖已有打印包：{output}（请换文件名或先删除）")

    sources = expand_sources(args.inputs)
    if not sources:
        raise SystemExit("没有匹配到任何文件")
    missing = [str(p) for p in sources if not p.exists()]
    if missing:
        raise SystemExit("找不到输入文件：" + "；".join(missing))

    copies = args.copies if args.copies is not None else default_copies(args.kind)
    if copies < 1:
        raise SystemExit(f"--kind {args.kind} 的份数为 0（辅助凭证不打印）")

    with tempfile.TemporaryDirectory(prefix="invoice_packet_") as tmp:
        items, notes = resolve_inputs(sources, tmp and Path(tmp))
        if not items:
            raise SystemExit("没有可用的发票 PDF")
        items, prune_notes = prune_supplements(items)
        if not items:
            raise SystemExit("剔除辅助凭证后没有可打印的发票")
        for note in notes + prune_notes:
            print(note)

        # 箱体处理：裁剪可见区到"内容 + 对称留白"。部分厂商发票的 CropBox 超出
        # MediaBox，或内容直接贴到纸边（实测某住宿发票右侧只剩 0.59mm），这两种
        # 情况在打印机不可打印边内都无法还原，必须先修 PDF 再排打印包。
        normalized: list[tuple[Path, Path | None]] = []
        for pdf, ofd in items:
            fixed, changed = normalize_boxes(pdf, Path(tmp), pad=16.0)
            if changed:
                print(f"{pdf.name}: 内容贴边或箱体异常，已裁到内容+16pt留白（源文件未改动）")
            normalized.append((fixed, ofd))
        items = normalized

        trip_date = args.trip_date or detect_trip_date(items) or datetime.now().strftime("%Y-%m-%d")

        source = fitz.open()
        for pdf, _ofd in items:
            with fitz.open(pdf) as doc:
                source.insert_pdf(doc)
        pages = [source[i] for i in range(len(source))] * max(1, copies)
        packet, per_sheet = build_invoice_packet(
            pages,
            trip_date=trip_date,
            per_sheet=args.per_sheet,
            margin=args.margin,
            gap=DEFAULT_GAP,
            cut_guides=args.cut_guides,
            copies=max(1, copies),
            number_base=args.start_number,
            number_total=args.total_number or None,
            slot_labels=[int(x) for x in args.slot_labels.split(",") if x.strip()] or None,
        )
        output.parent.mkdir(parents=True, exist_ok=True)
        packet.save(output)
        sheets = len(packet)
        packet.close()
        source.close()

    lines = [
        f"发票：{len(items)} 张 · 每张 {copies} 份 · 每张A4 {per_sheet} 张 · A4 共 {sheets} 页 · 标注日期 {trip_date}",
        f"打印包：{output}",
        "打印提示：用系统“适合页面/适合纸张”自动打印，不要手动缩放或居中。",
    ]
    for line in lines:
        print(line)
    if args.report:
        Path(args.report).expanduser().write_text("\n".join(lines) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
