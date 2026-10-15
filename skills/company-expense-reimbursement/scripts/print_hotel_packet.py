#!/usr/bin/env python
"""Hotel invoice print packet (住宿发票专用入口).

Thin wrapper over :mod:`print_invoice_packet` with the defaults confirmed on this
trip's hotel document:

  * kind = hotel, so the invoice is printed twice (lodging documents are filed
    twice; the 结账单 stays unprinted as a supporting document),
  * two invoices per A4 sheet, cropped to content, symmetric 10pt margins,
  * the label (check-in/check-out range + 第 x/y 页) sits in the blank strip below
    the invoice and never overlaps it,
  * content anchored top-left, never auto-centred, no printer-specific offsets,
  * abnormal PDF boxes are repaired automatically before printing,
  * 结账单/行程单 are dropped automatically: a lodging document is only kept when
    the invoice itself lacks the check-in or check-out date.

Usage:
    python scripts/print_hotel_packet.py <发票.pdf|目录|通配符> ... \
        --output-dir "05-报销" --name "住宿发票打印包.pdf"
"""

from __future__ import annotations

import sys

from print_invoice_packet import main as run_invoice_packet


def main() -> int:
    argv = sys.argv[1:]
    if not any(a == "--kind" or a.startswith("--kind=") for a in argv):
        argv = ["--kind", "hotel", *argv]
    sys.argv = [sys.argv[0], *argv]
    return run_invoice_packet()


if __name__ == "__main__":
    raise SystemExit(main())
