#!/usr/bin/env python
"""Report reimbursement totals without modifying finance-owned form fields."""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook

from reimbursement_amount import paper_total


def main() -> int:
    parser = argparse.ArgumentParser(description="Fill the paper reimbursement uppercase total.")
    parser.add_argument("workbook", help="Path to a reimbursement .xlsx workbook")
    args = parser.parse_args()
    path = Path(args.workbook).expanduser().resolve()
    workbook = load_workbook(path, data_only=True, read_only=True)
    total = 0
    for sheet in workbook.worksheets:
        if sheet.title.startswith("报销单"):
            total += paper_total(sheet)
    print(f"{path}\n非补助费用合计: {total:.2f}\n补助和大写金额由财务填写，工作簿未修改。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
