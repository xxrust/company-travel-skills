#!/usr/bin/env python
"""Create a new business-trip case folder with a minimal shared case record."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def yaml_scalar(value: str | None) -> str:
    if value is None or value == "":
        return "null"
    return json.dumps(value, ensure_ascii=False)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, help="New case directory")
    parser.add_argument("--case-id", required=True)
    parser.add_argument("--person", required=True)
    parser.add_argument("--department", default="")
    parser.add_argument("--purpose", default="")
    parser.add_argument("--start", default="")
    parser.add_argument("--end", default="")
    parser.add_argument("--origin", default="")
    args = parser.parse_args()

    root = Path(args.output).expanduser().resolve()
    if root.exists() and any(root.iterdir()):
        raise FileExistsError(f"Case directory is not empty: {root}")
    root.mkdir(parents=True, exist_ok=True)
    for name in ("01-出差申请", "02-出差过程", "03-发票凭证", "04-出差结束", "05-报销"):
        (root / name).mkdir(exist_ok=True)

    route = f"[{yaml_scalar(args.origin)}]" if args.origin else "[]"
    content = "\n".join(
        [
            f"case_id: {yaml_scalar(args.case_id)}",
            "status: draft",
            f"person: {yaml_scalar(args.person)}",
            f"department: {yaml_scalar(args.department)}",
            f"purpose: {yaml_scalar(args.purpose)}",
            f"planned_start: {yaml_scalar(args.start)}",
            f"planned_end: {yaml_scalar(args.end)}",
            f"planned_route: {route}",
            "actual_start: null",
            "actual_end: null",
            "actual_route: []",
            "budget: null",
            "approver: null",
            "application: {}",
            "daily_reports: []",
            "legs: []",
            "documents: []",
            "expenses: []",
            "open_issues: []",
            "events: []",
            "",
        ]
    )
    (root / "case.yaml").write_text(content, encoding="utf-8")
    (root / "issues.md").write_text("# 未解决问题\n\n暂无。\n", encoding="utf-8")
    print(root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
