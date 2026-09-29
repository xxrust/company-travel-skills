#!/usr/bin/env python
"""Run repository-level checks without reading or requiring private profile data."""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import yaml


ROOT = Path(__file__).resolve().parents[1]
SKILLS = ROOT / "skills"
NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


def read_frontmatter(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise ValueError(f"Missing YAML frontmatter: {path}")
    end = text.find("\n---", 4)
    if end < 0:
        raise ValueError(f"Unclosed YAML frontmatter: {path}")
    data = yaml.safe_load(text[4:end]) or {}
    if not isinstance(data, dict):
        raise ValueError(f"Frontmatter is not a mapping: {path}")
    if "TODO" in text or "[TODO" in text:
        raise ValueError(f"Unfinished scaffold marker: {path}")
    return data


def main() -> int:
    errors: list[str] = []
    skill_dirs = sorted(p for p in SKILLS.iterdir() if p.is_dir())
    if not skill_dirs:
        errors.append("No skill directories found")
    for skill_dir in skill_dirs:
        name = skill_dir.name
        skill_md = skill_dir / "SKILL.md"
        if not NAME_RE.fullmatch(name):
            errors.append(f"Invalid skill directory name: {name}")
        if not skill_md.exists():
            errors.append(f"Missing SKILL.md: {skill_dir}")
            continue
        try:
            frontmatter = read_frontmatter(skill_md)
            if frontmatter.get("name") != name:
                errors.append(f"Frontmatter name mismatch: {skill_md}")
            if not frontmatter.get("description"):
                errors.append(f"Missing description: {skill_md}")
        except Exception as exc:
            errors.append(str(exc))

    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1

    print(f"Validated {len(skill_dirs)} skills")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
