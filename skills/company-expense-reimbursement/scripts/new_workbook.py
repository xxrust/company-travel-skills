#!/usr/bin/env python
"""Copy the reusable reimbursement workbook template to a new claim file."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from reimbursement_profile import load_company_profile, load_user_profile


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, help="Destination .xlsx path")
    parser.add_argument("--profile", help="Local company profile YAML")
    parser.add_argument("--user-profile", help="Local traveler profile YAML")
    args = parser.parse_args()

    destination = Path(args.output).expanduser().resolve()
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("--output must end with .xlsx")
    skill_root = Path(__file__).resolve().parents[1]
    template = skill_root / "assets" / "company-expense-template.xlsx"
    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise FileExistsError(f"Refusing to overwrite existing workbook: {destination}")

    company, company_path = load_company_profile(args.profile)
    user, user_path = load_user_profile(args.user_profile)
    if company_path or user_path or args.profile or args.user_profile:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from build_template import make_workbook

        make_workbook(company, user).save(destination)
        print(f"{destination}\ncompany_profile={company_path or 'none'} user_profile={user_path or 'none'}")
    else:
        shutil.copy2(template, destination)
        print(f"{destination}\ncompany_profile=none user_profile=none")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
