#!/usr/bin/env python3
"""Quick sample regression — exit 1 if QA pack or web-style twins regress."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from main import extract_invoice, looks_like_part_number  # noqa: E402


def main() -> int:
    qa = ROOT.parent / "samples/RE_ Invoice PDF_Image to Excel Generator Tool-Status on 30-07-26"
    web = ROOT.parent / "samples/web-tests"
    fails = 0

    print("build check via import OK")
    if qa.is_dir():
        for pdf in sorted(qa.glob("*.pdf")):
            out = extract_invoice(pdf.read_bytes(), pdf.name)
            items = out.get("line_items") or []
            parts = sum(1 for it in items if looks_like_part_number(it.get("part_number") or ""))
            ok = len(items) >= 1 and out.get("supplier_name") not in ("", "Unknown", None)
            status = "OK" if ok else "FAIL"
            if not ok:
                fails += 1
            print(f"  {status} {pdf.name[:42]:42} n={len(items):3} parts={parts:3}")

    for name in ("web-ashok-style.pdf", "web-autolight-style.pdf", "web-alagu-style.pdf"):
        pdf = web / name
        if not pdf.exists():
            continue
        out = extract_invoice(pdf.read_bytes(), pdf.name)
        items = out.get("line_items") or []
        bad = [it.get("part_number") for it in items if (it.get("part_number") or "") in {"Wheel", "BOOT", "Minda"}]
        ok = len(items) >= 1 and not bad
        if not ok:
            fails += 1
        print(f"  {'OK' if ok else 'FAIL'} {name:42} n={len(items):3} bad_parts={bad}")

    print("FAILS", fails)
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
