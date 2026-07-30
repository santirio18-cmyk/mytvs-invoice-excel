"""
Detect invoice table column layout from header text — not vendor names.

Schemas drive which parsers run and whether MRP/Dis% post-repair is allowed.
"""

from __future__ import annotations

import re
from typing import Literal

LayoutSchema = Literal[
    "mrp_disc",
    "mrp_rate",
    "credit_rate_qty",
    "item_code",
    "einvoice",
    "unknown",
]

# Layouts that must never invent MRP from Rate
NO_MRP_REPAIR: frozenset[str] = frozenset(
    {"credit_rate_qty", "item_code", "einvoice"}
)

# Layouts that allow Ashok-style Qty/MRP/Dis% repair
MRP_REPAIR_OK: frozenset[str] = frozenset({"mrp_disc", "mrp_rate"})


def _header_blob(text: str, max_lines: int = 80) -> str:
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    return "\n".join(lines[:max_lines])


def detect_table_layout(text: str) -> LayoutSchema:
    """
    Infer table schema from column headers in the first ~80 lines.

    Priority (most specific first):
      item_code → mrp_disc → mrp_rate → credit_rate_qty → einvoice → unknown
    """
    hdr = _header_blob(text)
    if not hdr:
        return "unknown"

    has_mrp = bool(re.search(r"(?i)\bmrp\b", hdr))
    has_disc = bool(re.search(r"(?i)\bdis(?:c|count)?\s*%|\bdisc\s*%|\bdis\s*%", hdr))
    has_item_code = bool(re.search(r"(?i)item\s*code", hdr))
    has_particulars = bool(re.search(r"(?i)particulars", hdr))
    has_credit = bool(re.search(r"(?i)credit\s*bill", hdr))
    has_rate = bool(re.search(r"(?i)\brate\b", hdr))
    has_qty = bool(re.search(r"(?i)\bqty\b|\bquantity\b", hdr))
    has_hsn = bool(re.search(r"(?i)\bhsn\b", hdr))
    has_sno = bool(re.search(r"(?i)\b(?:s\.?\s*no|sl\.?\s*no|sr\.?\s*no)\b", hdr))
    has_part = bool(re.search(r"(?i)\bpart\s*(?:no|number|#|code)?\b|PARTNO", hdr))
    has_brand = bool(re.search(r"(?i)\bbrand\b", hdr))
    has_gstr = bool(re.search(r"(?i)gstr\s*%|gst\s*%", hdr))

    # Karnavati / ZipERP style
    if has_item_code and has_particulars:
        return "item_code"

    # PART No + PARTICULARS + GSTR% (ALA / Autolight Coimbatore) — treat as einvoice family
    if has_part and has_particulars and has_hsn:
        return "einvoice"

    # DESCRIPTION + PARTNO + BRAND (Karpagam)
    if has_brand and has_part and has_hsn:
        return "einvoice"

    # Item Rate + Unit Rate + Tax% (Vijayalakshmi / similar) — OCR may drop leading I
    if re.search(r"(?i)(?:item|\btem)\s*rate", hdr) and re.search(r"(?i)unit\s*rate", hdr):
        return "einvoice"

    # Ashok: Qty MRP Dis% Tax% Amount
    if has_mrp and has_disc:
        return "mrp_disc"

    # Spare parts with MRP + Rate columns
    if has_mrp and has_rate:
        return "mrp_rate"

    # Optech / Karthick digital e-invoice: S.No PartNo Description ...
    # (header may wrap "Tax %" across lines — don't require tax% on one line)
    if has_sno and has_part and has_hsn and re.search(r"(?i)\bdescription\b", hdr) and not has_mrp:
        return "einvoice"

    # Counter credit bill: Part HSN Rate Qty (no MRP column)
    if has_credit and not has_mrp:
        return "credit_rate_qty"

    # Same shape without the "CREDIT BILL" title: Part + HSN + Rate + Qty, no MRP
    if has_hsn and has_rate and has_qty and not has_mrp and (has_part or has_credit):
        # Prefer einvoice when S.No + tax% style is clear and amount follows rate
        if has_sno and re.search(r"(?i)\btax\s*%|\bgst\s*%", hdr):
            return "einvoice"
        return "credit_rate_qty"

    # Digital e-invoice: S.No Part Desc HSN Qty Rate Amount
    if has_sno and has_hsn and has_qty and has_rate and not has_mrp:
        return "einvoice"

    return "unknown"


def schema_allows_mrp_repair(schema: LayoutSchema, text: str = "") -> bool:
    """Whether Qty/MRP/Dis% post-repair may run for this schema."""
    if schema in NO_MRP_REPAIR:
        return False
    if schema in MRP_REPAIR_OK:
        return True
    # unknown: only if header clearly has MRP or Dis%
    hdr = _header_blob(text)
    return bool(
        re.search(r"(?i)\bmrp\b", hdr)
        or re.search(r"(?i)\bdis(?:c|count)?\s*%|\bdisc\s*%", hdr)
    )


def schema_expects_mrp(schema: LayoutSchema) -> bool:
    return schema in MRP_REPAIR_OK
