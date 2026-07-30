"""Layout-first regression: Ashok / credit-bill / Karnavati / e-invoice must not fight."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from extraction.layout import detect_table_layout, schema_allows_mrp_repair
from main import parse_line_items


ASHOK = """
TAX INVOICE
ASHOK AGENCIES
S.No Part No Description HSN Qty MRP Dis% Tax% Amount
1 ABC123456 WASHER SET 87089900 10 NOS 364.00 30.00 18.00 2548.00
2 DEF987654 BOLT KIT 87089900 5 NOS 200.00 10.00 18.00 900.00
"""

CREDIT = """
SRI PADMAVATHI AUTO CENTER
CREDIT BILL
Invoice No: SPAC/26-27/2498
Part No Description HSN Rate Qty
F8P08758 RADIATOR HOSES 87089900 247.00 3.00
B5431523 WIPER BLADE LH 85129000 383.00 7.00
F4915210 TONGUE WASHER 73182200 317.00 64.00
"""

KARNAVATI = """
KARNAVATI CAR AIRCONDITIONERS PRIVATE LIMITED
Bill To Order No. : Dated :
Sr. Item Code Particulars HSN Code Tax% Qty Rate Disc% Amount
No.
1 990030 COND SWIFT PTL/DSL T-1 84159000 18.00 5 PCS 1,700.00 8,500.00
2 991240 COND SWIFT DSL NEW T-2 84159000 18.00 5 PCS 1,925.00 9,625.00
"""

EINVOICE = """
TAX INVOICE
S.No Part No Description HSN Qty Rate Tax% Amount
1 1728431BC60SQMM25 8MM WIRE 85443000 2 NOS 1788.13 18 4219.98
2 B2K01702 ISOLATER SWITCH 85365000 2 NOS 762.71 18 1800.00
"""


def test_detect_schemas():
    assert detect_table_layout(ASHOK) == "mrp_disc"
    assert detect_table_layout(CREDIT) == "credit_rate_qty"
    assert detect_table_layout(KARNAVATI) == "item_code"
    assert detect_table_layout(EINVOICE) == "einvoice"


def test_mrp_repair_gating():
    assert schema_allows_mrp_repair("mrp_disc", ASHOK) is True
    assert schema_allows_mrp_repair("credit_rate_qty", CREDIT) is False
    assert schema_allows_mrp_repair("item_code", KARNAVATI) is False
    assert schema_allows_mrp_repair("einvoice", EINVOICE) is False


def test_ashok_mrp_filled():
    items = parse_line_items(ASHOK)
    assert items, "Ashok should yield rows"
    first = items[0]
    assert first.get("mrp"), f"MRP required: {first}"
    qty = float(str(first["qty"]).split()[0].replace(",", ""))
    assert qty == int(qty) and qty > 0
    assert float(str(first["rate"]).replace(",", "")) > 0


def test_credit_bill_no_invented_mrp():
    items = parse_line_items(CREDIT)
    assert len(items) >= 2
    first = items[0]
    assert first["part_number"] == "F8P08758"
    assert first["hsn_sac"] == "87089900"
    assert str(first["qty"]).startswith("3")
    assert first["rate"] == "247.00"
    assert "741" in first["amount"].replace(",", "")
    assert not first.get("mrp"), f"MRP must stay empty: {first}"


def test_karnavati_item_code():
    items = parse_line_items(KARNAVATI)
    assert len(items) >= 2
    first = items[0]
    assert first["part_number"] == "990030"
    assert "SWIFT" in first["description"].upper()
    assert "990030" not in first["description"]
    assert first["hsn_sac"] == "84159000"
    assert not first.get("mrp")


def test_einvoice_no_invented_mrp():
    items = parse_line_items(EINVOICE)
    assert items
    first = items[0]
    assert first.get("hsn_sac") == "85443000" or "85443000" in str(first)
    assert not first.get("mrp"), f"no invented MRP: {first}"


if __name__ == "__main__":
    test_detect_schemas()
    test_mrp_repair_gating()
    test_ashok_mrp_filled()
    test_credit_bill_no_invented_mrp()
    test_karnavati_item_code()
    test_einvoice_no_invented_mrp()
    print("ALL LAYOUT TESTS OK")
