#!/usr/bin/env python
"""Load local company and traveler profiles without embedding private data in the skill."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml


COMPANY_KEYS = ("name", "tax_id", "address", "phone", "bank_name", "bank_account")
USER_KEYS = ("name", "department", "title", "employee_id")


def _read_yaml(path: Path) -> dict[str, Any]:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError(f"Profile must contain a YAML mapping: {path}")
    return data


def _candidate_paths(explicit: str | None, env_name: str, filename: str) -> list[Path]:
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit).expanduser())
    env_value = os.environ.get(env_name)
    if env_value:
        candidates.append(Path(env_value).expanduser())
    candidates.append(Path.home() / ".codex" / "private" / filename)
    return candidates


def _find_profile(explicit: str | None, env_name: str, filename: str) -> Path | None:
    for candidate in _candidate_paths(explicit, env_name, filename):
        if candidate.exists():
            return candidate.resolve()
    return None


def load_company_profile(explicit: str | None = None, *, discover: bool = True) -> tuple[dict[str, str], Path | None]:
    path = _find_profile(explicit, "CODEX_EXPENSE_PROFILE", "company-profile.yaml") if discover else (Path(explicit).expanduser().resolve() if explicit else None)
    profile = {key: "" for key in COMPANY_KEYS}
    if path:
        data = _read_yaml(path)
        source = data.get("company", data)
        if not isinstance(source, dict):
            raise ValueError(f"Company profile must be a mapping: {path}")
        aliases = {
            "name": ("name", "company_name"),
            "tax_id": ("tax_id", "tax_number", "统一社会信用代码", "税号"),
            "address": ("address", "地址"),
            "phone": ("phone", "电话"),
            "bank_name": ("bank_name", "开户行"),
            "bank_account": ("bank_account", "account", "账号"),
        }
        for key, names in aliases.items():
            for name in names:
                if source.get(name) not in (None, ""):
                    profile[key] = str(source[name])
                    break
    return profile, path


def load_user_profile(explicit: str | None = None, *, discover: bool = True) -> tuple[dict[str, str], Path | None]:
    path = _find_profile(explicit, "CODEX_USER_PROFILE", "user-profile.yaml") if discover else (Path(explicit).expanduser().resolve() if explicit else None)
    profile = {key: "" for key in USER_KEYS}
    if path:
        data = _read_yaml(path)
        source = data.get("user", data)
        if not isinstance(source, dict):
            raise ValueError(f"User profile must be a mapping: {path}")
        aliases = {
            "name": ("name", "姓名"),
            "department": ("department", "部门"),
            "title": ("title", "职务", "职别"),
            "employee_id": ("employee_id", "工号"),
        }
        for key, names in aliases.items():
            for name in names:
                if source.get(name) not in (None, ""):
                    profile[key] = str(source[name])
                    break
    return profile, path
