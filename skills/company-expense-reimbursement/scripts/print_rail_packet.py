#!/usr/bin/env python
"""Rail e-ticket print packet (铁路发票专用入口).

Thin wrapper over :mod:`print_invoice_packet` with the railway defaults frozen
after the layout was confirmed in real printing:

  * kind = train, so each e-ticket is printed twice,
  * two e-tickets per A4 sheet, cropped to content, symmetric margins,
  * the label (travel-date range + 第 x/y 页) sits in the blank strip below the
    invoice and never overlaps it,
  * content anchored top-left, never auto-centred,
  * no printer-specific offsets or scaling — the printer's own "fit to page"
    does the final scaling, so a new printer needs no new parameters,
  * 行程单 is dropped automatically when the e-tickets already state the travel
    date and both stations.

Usage:
    python scripts/print_rail_packet.py <发票.pdf|发票.ofd|下载的.zip|目录> ... \
        --output "05-报销/铁路发票打印包.pdf"
"""

from __future__ import annotations

import sys

from print_invoice_packet import main as run_invoice_packet


def main() -> int:
    argv = sys.argv[1:]
    if not any(a == "--kind" or a.startswith("--kind=") for a in argv):
        argv = ["--kind", "train", *argv]
    sys.argv = [sys.argv[0], *argv]
    return run_invoice_packet()


if __name__ == "__main__":
    raise SystemExit(main())
