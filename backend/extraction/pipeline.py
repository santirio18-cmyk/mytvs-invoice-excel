"""
Extraction pipeline used by market-grade invoice apps:
1) Invoice-aware AI when configured (OpenAI vision / AWS Textract)
2) Layout/table reconstruction from OCR geometry
3) Legacy text parsers as fallback
4) Always validate totals + confidence
"""

from __future__ import annotations

import io
import os
from typing import Any, Callable

from PIL import Image

from extraction.validate import validate_invoice
from extraction.bbox_table import extract_items_from_image
from extraction import openai_invoice, textract_invoice


def _ext(name: str) -> str:
    return ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""


def _image_bytes_for_ai(data: bytes, filename: str) -> tuple[bytes, str] | None:
    ext = _ext(filename)
    if ext in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}:
        mime = "image/png" if ext == ".png" else "image/jpeg"
        if ext == ".webp":
            mime = "image/webp"
        return data, mime
    if ext == ".pdf":
        png = openai_invoice.pdf_first_page_png(data)
        if png:
            return png, "image/png"
    return None


def _score_items(items: list[dict[str, str]]) -> int:
    score = 0
    for it in items:
        if not (it.get("description") or it.get("part_number")):
            continue
        score += 1
        if it.get("qty"):
            score += 1
        if it.get("rate"):
            score += 1
        if it.get("mrp"):
            score += 1
        if it.get("amount"):
            score += 2
        if it.get("part_number"):
            score += 1
        if it.get("hsn_sac"):
            score += 1
    return score


def extract_invoice_v2(
    data: bytes,
    filename: str,
    *,
    legacy_extract_text: Callable[[bytes, str], str],
    legacy_parse: Callable[[str], dict[str, Any]],
) -> dict[str, Any]:
    """
    Run best available extractor. `legacy_parse(text)` should return the old
    extract_invoice dict fields (without needing to call extract_text again).
    """
    mode = (os.getenv("EXTRACTOR") or "auto").strip().lower()
    candidates: list[dict[str, Any]] = []

    img_payload = _image_bytes_for_ai(data, filename)

    # 1) OpenAI vision (highest accuracy in public benchmarks when keyed)
    if mode in {"auto", "openai"} and openai_invoice.available() and img_payload:
        try:
            ai = openai_invoice.extract_with_openai(img_payload[0], img_payload[1])
            if ai and _score_items(ai.get("line_items") or []) > 0:
                ai["filename"] = filename
                candidates.append(ai)
        except Exception as exc:  # noqa: BLE001
            candidates.append(
                {
                    "filename": filename,
                    "invoice_number": "Unknown",
                    "supplier_name": "Unknown",
                    "date": "Unknown",
                    "place_of_supply": "Unknown",
                    "line_items": [],
                    "extractor": "openai:error",
                    "warnings": [f"OpenAI extractor failed: {exc}"],
                }
            )

    # 2) AWS Textract AnalyzeExpense
    if mode in {"auto", "textract"} and textract_invoice.available():
        try:
            tx_bytes = img_payload[0] if img_payload else data
            tx = textract_invoice.extract_with_textract(tx_bytes)
            if tx and _score_items(tx.get("line_items") or []) > 0:
                tx["filename"] = filename
                candidates.append(tx)
        except Exception as exc:  # noqa: BLE001
            candidates.append(
                {
                    "filename": filename,
                    "invoice_number": "Unknown",
                    "supplier_name": "Unknown",
                    "date": "Unknown",
                    "place_of_supply": "Unknown",
                    "line_items": [],
                    "extractor": "textract:error",
                    "warnings": [f"Textract extractor failed: {exc}"],
                }
            )

    # 3) Legacy text parsers + bbox table boost
    text = legacy_extract_text(data, filename)
    legacy = legacy_parse(text)
    legacy["filename"] = filename
    legacy["extractor"] = legacy.get("extractor") or "tesseract_rules"
    legacy["raw_text_preview"] = text[:3500]

    # Geometry table pass on raster
    try:
        if img_payload:
            pil = Image.open(io.BytesIO(img_payload[0])).convert("RGB")
            bbox_items = extract_items_from_image(pil)
            if _score_items(bbox_items) > _score_items(legacy.get("line_items") or []):
                legacy = dict(legacy)
                legacy["line_items"] = bbox_items
                legacy["extractor"] = "tesseract_layout"
    except Exception:
        pass

    candidates.append(legacy)

    def rank(inv: dict[str, Any]) -> tuple:
        items = inv.get("line_items") or []
        return (
            _score_items(items),
            1 if inv.get("invoice_number") not in ("", "Unknown", None) else 0,
            1 if inv.get("supplier_name") not in ("", "Unknown", None) else 0,
            len(items),
        )

    best = max(candidates, key=rank)
    # Fill header gaps from legacy OCR text if AI missed a field
    for field in ("invoice_number", "supplier_name", "date", "place_of_supply"):
        if best.get(field) in (None, "", "Unknown") and legacy.get(field) not in (None, "", "Unknown"):
            best[field] = legacy[field]
    if not best.get("raw_text_preview"):
        best["raw_text_preview"] = text[:3500]

    return validate_invoice(best, text)
