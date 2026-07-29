"""
Layout-aware line extraction from Tesseract word boxes.

Commercial tools reconstruct tables from geometry (x/y), not only line text.
"""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

import pytesseract
from PIL import Image


_MONEY = re.compile(r"^[\d,]+\.\d{2}$")
_QTY = re.compile(r"^\d+(?:[.,]\d+)?$")
_HSN = re.compile(r"^\d{4,8}$")
_PART = re.compile(r"^(?:[A-Z]{1,4}\s*)?\d{3,6}[A-Za-z]?$", re.I)


def _cluster_rows(words: list[dict[str, Any]], y_tol: int = 12) -> list[list[dict[str, Any]]]:
    if not words:
        return []
    words = sorted(words, key=lambda w: (w["top"], w["left"]))
    rows: list[list[dict[str, Any]]] = []
    cur: list[dict[str, Any]] = []
    cur_y = None
    for w in words:
        if cur_y is None or abs(w["top"] - cur_y) <= y_tol:
            cur.append(w)
            cur_y = w["top"] if cur_y is None else (cur_y + w["top"]) // 2
        else:
            rows.append(sorted(cur, key=lambda x: x["left"]))
            cur = [w]
            cur_y = w["top"]
    if cur:
        rows.append(sorted(cur, key=lambda x: x["left"]))
    return rows


def extract_items_from_image(img: Image.Image) -> list[dict[str, str]]:
    """Return line items using word bounding boxes."""
    gray = img.convert("L")
    data = pytesseract.image_to_data(gray, lang="eng", config="--oem 3 --psm 6", output_type=pytesseract.Output.DICT)
    words: list[dict[str, Any]] = []
    n = len(data["text"])
    for i in range(n):
        txt = (data["text"][i] or "").strip()
        if not txt or int(data["conf"][i] or -1) < 20:
            continue
        words.append(
            {
                "text": txt,
                "left": int(data["left"][i]),
                "top": int(data["top"][i]),
                "width": int(data["width"][i]),
                "height": int(data["height"][i]),
            }
        )

    rows = _cluster_rows(words)
    items: list[dict[str, str]] = []
    for row in rows:
        texts = [w["text"] for w in row]
        joined = " ".join(texts)
        if re.search(r"(?i)hsn|taxable|cgst|sgst|total|invoice|gstin|description of goods", joined):
            if not any(_MONEY.match(t.replace(",", "")) or _MONEY.match(t) for t in texts):
                continue

        moneys = [t for t in texts if _MONEY.match(t) or _MONEY.match(t.replace(",", ""))]
        if len(moneys) < 1:
            continue
        # Need qty-like + rate/amount pattern
        amount = moneys[-1]
        rate = moneys[-2] if len(moneys) >= 2 else ""
        # qty: last integer-like token before moneys
        qty = ""
        unit = ""
        for t in texts:
            if re.fullmatch(r"(?i)nos|pcs|pkt|set|sets|roll", t):
                unit = t.lower()
        for t in reversed(texts):
            if _MONEY.match(t) or _MONEY.match(t.replace(",", "")):
                continue
            if _QTY.match(t) and not _HSN.match(t):
                qty = t
                break
        if not qty:
            continue

        hsn = ""
        for t in texts:
            if _HSN.match(t) and len(t) >= 4:
                hsn = t
                break

        # description / part = tokens left of hsn/qty/money
        skip = set(moneys)
        if qty:
            skip.add(qty)
        if hsn:
            skip.add(hsn)
        if unit:
            skip.add(unit)
        desc_tokens = []
        for t in texts:
            if t in skip:
                break
            if re.fullmatch(r"\d{1,3}", t) and not desc_tokens:
                continue  # serial
            desc_tokens.append(t)
        if len(desc_tokens) < 1:
            continue
        part = ""
        if desc_tokens and (_PART.match(desc_tokens[0]) or (re.search(r"\d", desc_tokens[0]) and re.search(r"[A-Za-z]", desc_tokens[0]))):
            part = desc_tokens[0]
            if len(desc_tokens) > 1 and re.fullmatch(r"(?i)[A-Z]{1,3}", desc_tokens[0]) and re.match(r"^\d", desc_tokens[1]):
                part = f"{desc_tokens[0]} {desc_tokens[1]}"
                desc = " ".join(desc_tokens[2:]) or part
            else:
                desc = " ".join(desc_tokens[1:]) or part
        else:
            # S 2007 style split across tokens
            if len(desc_tokens) >= 2 and re.fullmatch(r"(?i)s|\$", desc_tokens[0]) and re.match(r"^\d{3,5}", desc_tokens[1]):
                part = f"S {desc_tokens[1]}"
                desc = " ".join(desc_tokens[2:]) or part
            else:
                desc = " ".join(desc_tokens)

        items.append(
            {
                "part_number": part.replace("$", "S"),
                "description": desc.strip(" -|")[:120],
                "hsn_sac": hsn,
                "qty": f"{qty} {unit}".strip(),
                "rate": rate,
                "amount": amount,
            }
        )

    # soft dedupe identical OCR duplicates from multi-pass geometry noise only
    seen: set[tuple] = set()
    unique: list[dict[str, str]] = []
    for it in items:
        sig = (it.get("part_number"), it.get("qty"), it.get("rate"), it.get("amount"), it.get("description", "")[:20])
        if sig in seen:
            continue
        seen.add(sig)
        unique.append(it)
    return unique
