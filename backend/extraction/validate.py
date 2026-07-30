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
        r"(?i)(taxable\s*(?:value|amt|amount)?|sub\s*total|total\s*(?:amount)?|g\.?\s*total|grand\s*total|net\s*amount)"
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
        elif "net" in label:
            # Net amount is GST-inclusive — still useful as an upper bound later
            subtotal_labeled.append(v)
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


def _fnum(val: str | None) -> float | None:
    if not val:
        return None
    try:
        return float(str(val).split()[0].replace(",", ""))
    except Exception:
        return None


def strip_part_from_description(it: dict[str, str]) -> dict[str, str]:
    """Keep item code only in Part Number — remove duplicate from Description."""
    part = str(it.get("part_number") or "").strip()
    desc = str(it.get("description") or "").strip()
    it = dict(it)

    # Pull leading numeric Item Code into Part Number when missing
    if not part and desc:
        m = re.match(r"^(\d{4,8})\s+(.+)$", desc)
        if m:
            part = m.group(1)
            desc = m.group(2).strip()
            it["part_number"] = part
            it["description"] = desc
        else:
            # VT 1203001(Elec Oil...) or LX 3630KIT(KFK...)
            m = re.match(
                r"^([A-Z]{1,6}\s*\d{3,8}[A-Z0-9\-]*)\s*[\(]\s*(.+)$",
                desc,
                re.I,
            )
            if m:
                part = re.sub(r"\s+", " ", m.group(1)).strip()
                desc = m.group(2).rstrip(")").strip()
                it["part_number"] = part
                it["description"] = desc
            else:
                m = re.match(r"^([A-Z]{2,6}\d{2,5}[A-Z0-9]*)\s+(.+)$", desc, re.I)
                if m:
                    part = m.group(1)
                    desc = m.group(2).strip()
                    it["part_number"] = part
                    it["description"] = desc

    if not part or not desc:
        return it
    variants = {part, part.replace(" ", ""), re.sub(r"\s+", " ", part)}
    for v in list(variants):
        variants.add(v.upper())
        variants.add(v.lower())
    cleaned = desc
    for v in sorted(variants, key=len, reverse=True):
        if not v or len(v) < 3:
            continue
        pat = re.compile(rf"^{re.escape(v)}(\s*[-–:|/]\s*|\s+)", re.I)
        cleaned2 = pat.sub("", cleaned, count=1)
        if cleaned2 != cleaned:
            cleaned = cleaned2
            break
        if re.fullmatch(re.escape(v), cleaned, re.I):
            cleaned = ""
            break
    cleaned = cleaned.strip(" -–:|/")
    # Drop trailing HSN + tax% accidentally left in description
    cleaned = re.sub(r"\s+\d{4,8}\s+\d{1,2}(?:\.\d{1,2})?\s*$", "", cleaned).strip()
    it["description"] = cleaned
    return it


def _tidy_qty_only(items: list[dict[str, str]]) -> list[dict[str, str]]:
    """Integer-ize whole qtys; clear fake MRP==Rate copies. No inventing MRP."""
    fixed: list[dict[str, str]] = []
    for raw in items:
        it = dict(raw)
        qty_s = str(it.get("qty") or "").strip()
        qty_tok = qty_s.split()[0].replace(",", "") if qty_s else ""
        qty_v = _fnum(qty_tok)
        rate_v = _fnum(it.get("rate"))
        mrp_v = _fnum(it.get("mrp"))
        if qty_v is not None and qty_v == int(qty_v):
            unit = " ".join(qty_s.split()[1:]) if qty_s.split()[1:] else ""
            it["qty"] = f"{int(qty_v)} {unit}".strip()
        if mrp_v is not None and rate_v is not None and abs(mrp_v - rate_v) < 0.01:
            it["mrp"] = ""
        fixed.append(it)
    return fixed


def repair_qty_mrp_shift(
    items: list[dict[str, str]],
    text: str = "",
    schema: str | None = None,
) -> list[dict[str, str]]:
    """
    Stop MRP / discount landing in Qty (Ashok Qty→MRP→Dis%→Tax%→Amount layouts).

    Schema-gated (layout-first):
      - credit_rate_qty / item_code / einvoice → tidy only, never invent MRP
      - mrp_disc / mrp_rate → full repair
      - unknown → repair only if header has MRP or Dis%
    """
    from extraction.layout import detect_table_layout, schema_allows_mrp_repair

    resolved = schema or detect_table_layout(text or "")
    if not schema_allows_mrp_repair(resolved, text or ""):  # type: ignore[arg-type]
        return _tidy_qty_only(items)

    has_mrp_hdr = bool(re.search(r"(?i)\bmrp\b", (text or "")[:2000]))
    fixed: list[dict[str, str]] = []
    for raw in items:
        it = dict(raw)
        qty_s = str(it.get("qty") or "").strip()
        qty_tok = qty_s.split()[0].replace(",", "") if qty_s else ""
        mrp_s = str(it.get("mrp") or "").strip()
        rate_s = str(it.get("rate") or "").strip()
        amt_s = str(it.get("amount") or "").strip()
        qty_v = _fnum(qty_tok)
        mrp_v = _fnum(mrp_s)
        rate_v = _fnum(rate_s)
        amt_v = _fnum(amt_s)

        def _set_qty(q: float) -> None:
            it["qty"] = str(int(round(q))) if abs(q - round(q)) < 0.08 else f"{q:.2f}"

        def _derive_qty_from_amount(unit: float) -> bool:
            if not amt_v or not unit or unit <= 0:
                return False
            q = amt_v / unit
            if 0.5 <= q <= 5000:
                _set_qty(q)
                return True
            return False

        # Already a clean Qty × Rate ≈ Amount line — do not invent MRP/discount
        line_balanced = (
            qty_v is not None
            and rate_v is not None
            and amt_v is not None
            and qty_v > 0
            and abs(qty_v * rate_v - amt_v) <= max(0.51, 0.02 * amt_v)
        )
        if line_balanced and (not mrp_v or abs((mrp_v or 0) - rate_v) < 0.01):
            if qty_v == int(qty_v):
                unit = " ".join(qty_s.split()[1:]) if qty_s.split()[1:] else ""
                it["qty"] = f"{int(qty_v)} {unit}".strip()
            if mrp_v and abs(mrp_v - rate_v) < 0.01 and not has_mrp_hdr:
                it["mrp"] = ""
            fixed.append(it)
            continue

        # C: Dis% in Qty, MRP in Rate
        if (
            not mrp_v
            and qty_v is not None
            and rate_v is not None
            and amt_v is not None
            and 0 < qty_v <= 100
            and re.search(r"\.\d{2}", qty_tok or "")
            and rate_v >= 50
            and amt_v > rate_v * 0.15
            and abs(qty_v * rate_v - amt_v) > max(1.0, 0.05 * amt_v)
        ):
            disc = qty_v
            it["mrp"] = rate_s
            net = rate_v * (1 - disc / 100.0)
            if net > 0 and _derive_qty_from_amount(net):
                it["rate"] = f"{net:.2f}"
            else:
                it["qty"] = ""
                it["rate"] = ""

        # A) Qty looks like money → it is MRP
        elif qty_tok and re.fullmatch(r"\d+\.\d{2}", qty_tok) and qty_v is not None and qty_v >= 100:
            if not mrp_v:
                it["mrp"] = qty_tok
                mrp_v = qty_v
            it["qty"] = ""
            unit = rate_v if rate_v and mrp_v and abs(rate_v - mrp_v) > 0.5 else None
            if unit and _derive_qty_from_amount(unit):
                pass
            elif mrp_v and amt_v and rate_v and 0 < rate_v <= 100:
                net = mrp_v * (1 - rate_v / 100.0)
                if net > 0 and _derive_qty_from_amount(net):
                    it["rate"] = f"{net:.2f}"

        # B) Qty is truncated MRP under an MRP header
        elif (
            has_mrp_hdr
            and qty_v is not None
            and qty_v >= 50
            and (rate_v is None or rate_v < 5)
            and not mrp_v
        ):
            it["mrp"] = f"{qty_v:.2f}"
            it["qty"] = ""
            if rate_v and rate_v > 0 and _derive_qty_from_amount(rate_v):
                pass

        # D) Rate was copied from MRP
        mrp_v = _fnum(it.get("mrp"))
        rate_v = _fnum(it.get("rate"))
        qty_v = _fnum(str(it.get("qty") or "").split()[0] if it.get("qty") else None)
        amt_v = _fnum(it.get("amount"))
        if (
            mrp_v
            and rate_v
            and abs(mrp_v - rate_v) < 0.01
            and qty_v
            and amt_v
            and qty_v > 0
        ):
            net = amt_v / qty_v
            if abs(net - mrp_v) > 0.5:
                it["rate"] = f"{net:.2f}"

        # E) Never leave a large money-looking value in Qty
        qty_s2 = str(it.get("qty") or "").strip()
        qty_tok2 = qty_s2.split()[0].replace(",", "") if qty_s2 else ""
        if qty_tok2 and re.fullmatch(r"\d+\.\d{2}", qty_tok2):
            q2 = _fnum(qty_tok2)
            if q2 is not None and q2 >= 100:
                if not it.get("mrp"):
                    it["mrp"] = qty_tok2
                it["qty"] = ""

        fixed.append(it)
    return fixed


def validate_invoice(inv: dict[str, Any], text: str = "") -> dict[str, Any]:
    """
    Attach confidence + warnings. Market tools always reconcile line sums to totals.
    """
    items = repair_qty_mrp_shift(
        list(inv.get("line_items") or []),
        text or inv.get("raw_text_preview") or "",
    )
    items = [strip_part_from_description(dict(it)) for it in items]
    inv = dict(inv)
    inv["line_items"] = items

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
            # Low-res screenshots often invent a few wrong rows — drop them
            # when the sum is nowhere near the invoice total.
            if taxable >= 500 and item_sum < taxable * 0.25:
                warnings.append(
                    "Line items cleared — OCR totals were unreliable. "
                    "Re-upload a clearer PDF/photo (or enable OpenAI on the server)."
                )
                items = []
                inv = dict(inv)
                inv["line_items"] = items
                item_sum = 0.0
                score = min(score, 0.4)

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
