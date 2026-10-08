#!/usr/bin/env python
"""Convert a reimbursement paper-form HTML to a print-ready PDF via headless Chromium."""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path


def html_to_pdf(html_path: Path, output: Path, *, margins_mm: int = 5) -> dict[str, object]:
    """Render *html_path* to *output* using Playwright headless Chromium.

    Returns a manifest dict with paths, page count, and timestamp.
    """
    from playwright.sync_api import sync_playwright  # imported here so the module

    # is importable even when playwright is not installed (fail at call time).
    output.parent.mkdir(parents=True, exist_ok=True)
    margin = f"{margins_mm}mm"

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(html_path.as_uri(), wait_until="networkidle")
        page.pdf(
            path=str(output),
            format="A4",
            print_background=True,
            margin={"top": margin, "bottom": margin, "left": margin, "right": margin},
        )
        browser.close()

    # Count pages via pymupdf / fitz (already a project dependency via print_packet).
    page_count: int | None = None
    try:
        import fitz  # type: ignore
        doc = fitz.open(str(output))
        page_count = len(doc)
        doc.close()
    except Exception:
        pass

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_html": str(html_path),
        "output_pdf": str(output),
        "page_count": page_count,
    }
    manifest_path = output.with_suffix(".json")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert reimbursement HTML to print-ready PDF via headless Chromium.")
    parser.add_argument("html", help="Path to the paper-form HTML file")
    parser.add_argument("--output", required=True, help="Destination PDF path")
    parser.add_argument("--margins", type=int, default=5, help="Page margin in mm (default: 5)")
    parser.add_argument("--print", action="store_true", dest="do_print", help="Send to default printer after generating")
    parser.add_argument("--printer", help="Named printer (requires --print)")
    args = parser.parse_args()

    html_path = Path(args.html).expanduser().resolve()
    if not html_path.exists():
        raise FileNotFoundError(f"HTML file not found: {html_path}")

    output = Path(args.output).expanduser().resolve()
    manifest = html_to_pdf(html_path, output, margins_mm=args.margins)
    print(json.dumps(manifest, ensure_ascii=False, indent=2))

    if args.do_print:
        if args.printer:
            os.startfile(str(output), "printto", f'"{args.printer}"')
        else:
            os.startfile(str(output), "print")
        print(f"已发送至打印机：{args.printer or '默认'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
