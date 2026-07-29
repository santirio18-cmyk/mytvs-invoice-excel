"""Optional AWS Textract AnalyzeExpense — purpose-built invoice/receipt API."""

from __future__ import annotations

import os
from typing import Any


def available() -> bool:
    if os.getenv("AWS_ACCESS_KEY_ID", "").strip() and os.getenv("AWS_SECRET_ACCESS_KEY", "").strip():
        return True
    return os.getenv("TEXTRACT_ENABLED", "").lower() in {"1", "true", "yes"}


def extract_with_textract(document_bytes: bytes) -> dict[str, Any] | None:
    if not available():
        # Still try if boto3 can resolve credentials another way and flag is on
        if os.getenv("TEXTRACT_ENABLED", "").lower() not in {"1", "true", "yes"}:
            if not (os.getenv("AWS_ACCESS_KEY_ID") and os.getenv("AWS_SECRET_ACCESS_KEY")):
                return None
    try:
        import boto3
    except Exception:
        return None

    region = os.getenv("AWS_REGION") or os.getenv("AWS_DEFAULT_REGION") or "ap-south-1"
    client = boto3.client("textract", region_name=region)
    resp = client.analyze_expense(Document={"Bytes": document_bytes})

    inv_no = supplier = date = place = "Unknown"
    items: list[dict[str, str]] = []

    for doc in resp.get("ExpenseDocuments") or []:
        for field in doc.get("SummaryFields") or []:
            ftype = ((field.get("Type") or {}).get("Text") or "").upper()
            val = ((field.get("ValueDetection") or {}).get("Text") or "").strip()
            if not val:
                continue
            if ftype in {"INVOICE_RECEIPT_ID", "INVOICE_ID"}:
                inv_no = val
            elif ftype in {"VENDOR_NAME", "SUPPLIER_NAME", "NAME"}:
                if supplier == "Unknown":
                    supplier = val
            elif ftype in {"INVOICE_RECEIPT_DATE", "DATE"}:
                date = val
            elif "ADDRESS" in ftype and place == "Unknown":
                place = val[:80]

        for group in doc.get("LineItemGroups") or []:
            for line in group.get("LineItems") or []:
                row = {
                    "part_number": "",
                    "description": "",
                    "hsn_sac": "",
                    "qty": "",
                    "rate": "",
                    "amount": "",
                }
                for f in line.get("LineItemExpenseFields") or []:
                    ftype = ((f.get("Type") or {}).get("Text") or "").upper()
                    val = ((f.get("ValueDetection") or {}).get("Text") or "").strip()
                    if not val:
                        continue
                    if ftype in {"ITEM", "PRODUCT_CODE"}:
                        if ftype == "PRODUCT_CODE" or not row["part_number"]:
                            if ftype == "PRODUCT_CODE":
                                row["part_number"] = val
                            else:
                                row["description"] = val
                        else:
                            row["description"] = val
                    elif ftype == "QUANTITY":
                        row["qty"] = val
                    elif ftype in {"UNIT_PRICE", "PRICE"}:
                        row["rate"] = val
                    elif ftype in {"PRICE", "AMOUNT", "EXPENSE_ROW"}:
                        if ftype in {"AMOUNT", "EXPENSE_ROW"} or not row["amount"]:
                            row["amount"] = val
                    elif "HSN" in ftype or ftype == "PRODUCT_CODE":
                        if val.isdigit() and len(val) >= 4:
                            row["hsn_sac"] = val
                if row["description"] or row["part_number"] or row["amount"]:
                    items.append(row)

    if not items and inv_no == "Unknown" and supplier == "Unknown":
        return None

    return {
        "invoice_number": inv_no,
        "supplier_name": supplier,
        "date": date,
        "place_of_supply": place,
        "line_items": items,
        "extractor": "aws_textract",
    }
