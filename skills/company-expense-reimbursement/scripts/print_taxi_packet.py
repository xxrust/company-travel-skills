#!/usr/bin/env python
"""Ride-hailing invoice print packet (滴滴/网约车发票专用入口).

Thin wrapper over :mod:`print_invoice_packet` with the defaults frozen after the
layout and box handling were confirmed by the user:

  * kind = taxi, so each e-invoice is printed once,
  * two e-invoices per A4 sheet, cropped to content, symmetric 10pt margins,
  * the label (travel-date range + 第 x/y 页) sits in the blank strip below the
    invoice and never overlaps it,
  * content anchored top-left, never auto-centred, no printer-specific offsets —
    the printer's own "fit to page" does the final scaling,
  * abnormal PDF boxes are repaired automatically: some vendor e-invoices ship a
    CropBox wider than the MediaBox (610 vs 595.28, starting at y=-22), and every
    standards-compliant viewer then clips the intersection, silently cutting the
    left QR code and the right-hand invoice number. The original file is never
    modified; the packet uses a corrected copy,
  * 行程单 is dropped automatically when the e-invoices already state the travel
    date and both endpoints (they do for Didi invoices: 出行日期/出发地/到达地).

Usage:
    python scripts/print_taxi_packet.py <发票.pdf|目录|通配符> ... \
        --output-dir "05-报销" --name "滴滴发票打印包.pdf"
"""

from __future__ import annotations

import sys

from print_invoice_packet import main as run_invoice_packet


def main() -> int:
    argv = sys.argv[1:]
    if not any(a == "--kind" or a.startswith("--kind=") for a in argv):
        argv = ["--kind", "taxi", *argv]
    if not any(a == "--copies" or a.startswith("--copies=") for a in argv):
        argv = ["--copies", "1", *argv]
    sys.argv = [sys.argv[0], *argv]
    return run_invoice_packet()


if __name__ == "__main__":
    raise SystemExit(main())
