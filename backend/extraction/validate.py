"""Post-extraction checks used by commercial invoice products."""

from __future__ import annotations

import re
from typing import Any


def _money(val: str | None) -> float | None:
    if not val:
        return None
    t = str(val).strip().replace(",", "")
    try:
        return float(t)
    except Exception:
        return None


def find_taxable_total(text: str, item_sum: float | None = None) -> float | None:
    """Pick the most likely taxable / subtotal from OCR text (not GST grand total)."""
    taxable_labeled: list[float] = []
    subtotal_labeled: list[float] = []
    for m in re.finditer(
        r"(?i)(taxable\s*(?:value|amt|amount)?|sub\s*total|total\s*(?:amount)?|g\.?\s*total|grand\s*total)"
        r"\s*[:\-]?\s*([\d,]+\.\d{2})",
        text or "",
    ):
        label = m.group(1).lower()
        v = _money(m.group(2))
        if not v or not (10 <= v <= 5_000_000):
            continue
        if "taxable" in label or "sub" in label:
            taxable_labeled.append(v)
        elif "grand" in label or "g." in label or "g total" in label:
            continue  # skip GST-inclusive grand total
        else:
            subtotal_labeled.append(v)

    solos: list[float] = []
    for raw in (text or "").splitlines():
        ln = raw.strip().replace("|", " ")
        if re.fullmatch(r"[\d,]+\.\d{2}", ln):
            v = _money(ln)
            if v and 500 <= v <= 5_000_000:
                solos.append(v)

    pool = taxable_labeled or subtotal_labeled or solos
    if not pool:
        return None
    if item_sum and item_sum > 0:
        # Prefer value closest to line-item sum (taxable), not grand total
        return min(pool, key=lambda x: abs(x - item_sum))
    pool.sort()
    return pool[len(pool) // 2]


def validate_invoice(inv: dict[str, Any], text: str = "") -> dict[str, Any]:
    """
    Attach confidence + warnings. Market tools always reconcile line sums to totals.
    """
    items = list(inv.get("line_items") or [])
    warnings: list[str] = list(inv.get("warnings") or [])
    amounts = [_money(it.get("amount")) for it in items]
    amounts_f = [a for a in amounts if a is not None and a > 0]
    item_sum = round(sum(amounts_f), 2) if amounts_f else 0.0
    taxable = find_taxable_total(text or inv.get("raw_text_preview") or "", item_sum=item_sum)

    score = 0.35
    if inv.get("invoice_number") not in (None, "", "Unknown"):
        score += 0.15
    if inv.get("supplier_name") not in (None, "", "Unknown"):
        score += 0.1
    if inv.get("date") not in (None, "", "Unknown"):
        score += 0.05
    if items:
        score += min(0.25, 0.04 * len(items))

    if not items:
        warnings.append("No line items detected — re-upload a clearer PDF/scan.")
        score = min(score, 0.35)
    elif taxable is not None:
        diff = round(taxable - item_sum, 2)
        if abs(diff) <= 1.0:
            score += 0.15
        elif abs(diff) <= max(5.0, taxable * 0.02):
            warnings.append(
                f"Line items sum ₹{item_sum:,.2f} vs taxable ≈ ₹{taxable:,.2f} "
                f"(₹{diff:+.2f}) — review before posting."
            )
            score += 0.05
        else:
            warnings.append(
                f"Totals do not match: lines ₹{item_sum:,.2f} vs taxable ₹{taxable:,.2f} "
                f"(₹{diff:+.2f}). Check missing/extra rows."
            )
            score -= 0.1

    missing_amt = sum(1 for it in items if not _money(it.get("amount")))
    if missing_amt:
        warnings.append(f"{missing_amt} line(s) missing amount.")
        score -= 0.05

    score = max(0.05, min(0.99, score))
    if score >= 0.8:
        confidence = "high"
    elif score >= 0.55:
        confidence = "medium"
    else:
        confidence = "low"

    inv = dict(inv)
    inv["warnings"] = warnings
    inv["confidence"] = confidence
    inv["confidence_score"] = round(score, 2)
    inv["items_sum"] = item_sum
    if taxable is not None:
        inv["taxable_total"] = taxable
    return inv
