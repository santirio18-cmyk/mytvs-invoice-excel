"""Optional OpenAI vision invoice extraction (market-leading multimodal approach)."""

from __future__ import annotations

import json
import os
import re
from typing import Any


SYSTEM = """You extract Indian GST tax invoice data for automotive spare parts.
Return ONLY valid JSON with this shape:
{
  "invoice_number": string,
  "supplier_name": string,
  "date": string,
  "place_of_supply": string,
  "taxable_total": string,
  "line_items": [
    {"part_number": string, "description": string, "hsn_sac": string, "qty": string, "mrp": string, "rate": string, "amount": string}
  ]
}
Rules:
- Include EVERY line item row, including duplicate SKUs listed twice.
- qty may include unit (e.g. "50 nos").
- If the invoice has an MRP column, fill mrp. When there is no separate Rate column
  (Qty → MRP → Dis% → Tax% → Amount), put the MRP in mrp and set rate to Amount/Qty
  (net unit price), not the discount percent.
- Prefer seller/supplier name, not the buyer (TVS / consignee).
- If unsure, use empty string — never invent HSN/amounts.
"""


def available() -> bool:
    return bool(os.getenv("OPENAI_API_KEY", "").strip())


def extract_with_openai(image_bytes: bytes, mime: str = "image/png") -> dict[str, Any] | None:
    if not available():
        return None
    try:
        from openai import OpenAI
    except Exception:
        return None

    import base64

    model = os.getenv("OPENAI_INVOICE_MODEL", "gpt-4o").strip() or "gpt-4o"
    b64 = base64.b64encode(image_bytes).decode("ascii")
    client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])
    resp = client.chat.completions.create(
        model=model,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": SYSTEM},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Extract all invoice fields and every line item."},
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ],
            },
        ],
    )
    raw = resp.choices[0].message.content or "{}"
    data = json.loads(raw)
    items = []
    for it in data.get("line_items") or []:
        if not isinstance(it, dict):
            continue
        items.append(
            {
                "part_number": str(it.get("part_number") or "").strip(),
                "description": str(it.get("description") or "").strip(),
                "hsn_sac": str(it.get("hsn_sac") or "").strip(),
                "qty": str(it.get("qty") or "").strip(),
                "mrp": str(it.get("mrp") or "").strip(),
                "rate": str(it.get("rate") or "").strip(),
                "amount": str(it.get("amount") or "").strip(),
            }
        )
    return {
        "invoice_number": str(data.get("invoice_number") or "Unknown").strip() or "Unknown",
        "supplier_name": str(data.get("supplier_name") or "Unknown").strip() or "Unknown",
        "date": str(data.get("date") or "Unknown").strip() or "Unknown",
        "place_of_supply": str(data.get("place_of_supply") or "Unknown").strip() or "Unknown",
        "line_items": items,
        "extractor": f"openai:{model}",
        "taxable_hint": str(data.get("taxable_total") or ""),
    }


def pdf_first_page_png(data: bytes) -> bytes | None:
    try:
        import pypdfium2 as pdfium
        from io import BytesIO
        from PIL import Image

        doc = pdfium.PdfDocument(data)
        if len(doc) < 1:
            return None
        pil = doc[0].render(scale=2.0).to_pil().convert("RGB")
        buf = BytesIO()
        pil.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None
