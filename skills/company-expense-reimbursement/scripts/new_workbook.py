#!/usr/bin/env python
"""Build a new reimbursement workbook without overwriting existing claims."""

from __future__ import annotations

import argparse
from pathlib import Path

from reimbursement_profile import load_company_profile, load_user_profile
from build_template import make_workbook


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, help="Destination .xlsx path")
    parser.add_argument("--profile", help="Local company profile YAML")
    parser.add_argument("--user-profile", help="Local traveler profile YAML")
    args = parser.parse_args()

    destination = Path(args.output).expanduser().resolve()
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("--output must end with .xlsx")
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise FileExistsError(f"Refusing to overwrite existing workbook: {destination}")

    company, company_path = load_company_profile(args.profile)
    user, user_path = load_user_profile(args.user_profile)
    make_workbook(company, user).save(destination)
    print(f"{destination}\ncompany_profile={company_path or 'none'} user_profile={user_path or 'none'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
