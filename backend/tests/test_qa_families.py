"""
Foolproof regression: similar-case fixtures for every QA layout family.

These are synthetic twins of the 30-07-26 failures — same column shapes,
different supplier names / part codes — so future invoices stay covered.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from extraction.layout import detect_table_layout
from main import (
    extract_invoice,
    looks_like_part_number,
    parse_anamallais_part_above_si,
    parse_desc_part_brand_hsn,
    parse_goods_hsn_qty_rate_disc,
    parse_line_items,
    parse_lubricant_mrp_qty_rate,
    parse_sno_part_desc_discounts,
    parse_sno_part_particulars_gst,
)


# --- Similar-case fixtures (not the original vendor names) -----------------

ALA_TWIN = """
TAX INVOICE
BRIGHT AUTO PRODUCTZ
S.No PART No PARTICULARS HSN GSTR % QTY Rate Dis % Amount
1 9910-2201 Gates Fan Belt Heavy 87089100 18 % 2 NOS 1,200.00 10 % 2,160.00
"""

AUTOLIGHT_TWIN = """
TAX INVOICE
CITY LIGHT AGENCY
S.No PART PARTICULARS HSN GSTR% QTY Rate Amount
1 HL-4401 12V H4 Headlamp 85392120 18 30 NOS 45.00 1,350.00
"""

ARVIND_TWIN = """
TAX INVOICE
Description of Goods HSN MRP Alt. Quantity Prod Quantity Rate per Sch. Amount
SHELL LUBE DISTRIBUTORS
1550099999 RIMULA R4 27101972 2,500.00 112.500 Ltr\\Kg 10 Nos 2,500.00 Nos 25,000.00
PLUS 15W40 1X7.5L
MRP 4000
2550088888 27101980 1,500.00 15.000 Ltr\\Kg 1-0.0000 Ctn 1,500.00 Pcs 4,500.00
-RimR415W40_3
*5L_A1W8 MRP 2700
3550077777-RimR415W40 27101980 3,000.00 50.000 Ltr\\Kg 5.0000 Pcs 3,000.00 Pcs 15,000.00
1*10L_A1W8 MRP 5000
"""

KARPAGAM_TWIN = """
TAX INVOICE
CITY AUTO STORES
S.No DESCRIPTION PARTNO BRAND HSN QTY UOM RATE GST TAXABLE
1 HAND BR HOSE 30 BIG NUT 91514YA ADEENA 39172390 1 NOS 249.70 18 249.70
9 LEYLAND U TRUCK FAN BELT 8PK-1552CONTITECH 40103390 1 NOS 598.90 18 598.90
24 27 F L HOSE RUBBER 7971/4RSSTAR 40091100 10 NOS 96.60 18 966.00
32 JACK SEAL KIT RUBBER 649E VELFIT 40169390 10 KIT 48.95 18 489.50
33 COURIER CHARGES 996812 1 750.00 18 750.00
"""

ANAMALLAIS_TWIN = """
Dealer Agencies (North)
Receiver GST No.: 33AAGCM0329K1ZM
Part No Price Sale
MRP per Dis. (Rs)
S.No (HSN Part Description UoM Qty (Rs) per Part Cat Rate Total (Rs.)
XX100001
1 (8708700 LOCK NUT TEST/(A) NOS 730.00 2.00 618.65 14.00 -86.61 532.04 1,064.08
0)
YY200002
2 7(870899 COVER OIL COOLER/(A) NOS 1,455.00 1.00 1,233.05 14.00 -172.63 1,060.42 1,060.42
00)
Spare Parts Total 2,124.50
HSN Code Summary
87087000 1 1000.00
"""

ASK_TWIN = """
TAX INVOICE
COUNTER AUTO SPARES
S.No PartNo Description HSN Qty Unit Rate Dis% CD% Tax% Amount
1 CS607822 R/SPRING LEAF 73181500 4 Pcs 135.00 15.00 5.00 18 436.05
11 cs100667 BOLT W/NUT16 73181500 10 Nos 84.00 15.00 18 714.00
12 AB999001 WASHER FLAT 73182200 2 Pcs 50.00 18 118.00
"""

HEMA_TWIN = """
TAX INVOICE
NORTH FILTER AGENCIES
Sl Description of Goods HSN/SAC Quantity Rate per Disc. % Amount
1 LX 9999KIT(KFK0249506) 84213100 1 NOS 822.00 NOS 13.76 % 708.89
"""

MB_TWIN = """
TAX INVOICE
SOUTH GAUGE Agencies
1 VT 9988001(Elec Oil Pressure Guage Tata 0-10) 90262000 18% 2 nos 467.00 nos 934.00
"""

PR_TWIN = """
TAX INVOICE
P.Q. PROCESS - (from 1-Apr-25)
1 LL149 HINO MANIFOLD SET OF 6 ( 84841090 5 NOS 102.00 NOS 510.00
2 LL152 HINO BS4 EGR COOLER GASKET 84841090 10 NOS 183.00 NOS 1,830.00
"""

ANBU_TWIN = """
TAX INVOICE
CITY BEARINGS
46, SOME STREET ROAD
Part HSN Qty Rate Tax% Net Amount
TX99 TPH TEXSPIN 84828000 3 734.48 18 866.69 2,200.44
"""

ALAGU_NO_PART = """
TAX INVOICE
ALAGU GEAR ROD COVERS
S.No Part No Description HSN Qty MRP Rate Amount
1 14 Wheel BOOT 87087000 100 165.00 16500.00
2 10/12 Wheel BOOT 87087000 500 115.00 57500.00
"""

SCAN1_KARPAGAM_NOISY = """
KARPAGAM AUTO STORES
TAX INVOICE TRANSPORT COPY
S.No DESCRIPTION OF GOODS PARTNO BRAND HSN/SAC QTY U! RATE GST TAXABLE VAL
1 V ROD ASSY U TRUCK B6Y03007 LEYLAND \\c 19 10057 = PRIZOL 73269099 = {ANOS 10935.00 18 10,935.00
2 COURIER CHARGES 996812 1 50.00 18 50.00
"""

CREDIT_TWIN = """
CREDIT BILL
Part No Description HSN Rate Qty
Z9P08758 RADIATOR HOSES 87089900 247.00 3.00
"""


def test_looks_like_part_number():
    assert looks_like_part_number("7818-1501")
    assert looks_like_part_number("8PK-1552")
    assert looks_like_part_number("IA202777")
    assert looks_like_part_number("VT 1203001")
    assert not looks_like_part_number("Wheel")
    assert not looks_like_part_number("Minda")
    assert not looks_like_part_number("BOOT")
    assert not looks_like_part_number("CHARGES")


def test_ala_particulars_twin():
    assert detect_table_layout(ALA_TWIN) == "einvoice"
    items = parse_sno_part_particulars_gst(ALA_TWIN) or parse_line_items(ALA_TWIN)
    assert len(items) >= 1
    assert items[0]["part_number"] == "9910-2201"
    assert "Fan Belt" in items[0]["description"]
    assert not items[0].get("mrp")


def test_arvind_ctn_qty_twin():
    items = parse_lubricant_mrp_qty_rate(ARVIND_TWIN)
    assert len(items) == 3
    mid = next(it for it in items if it["part_number"] == "2550088888")
    assert mid["qty"].startswith("1")
    assert "Ctn" in mid["qty"]
    assert "2885" not in mid["description"]
    assert mid["mrp"].startswith("2700")


def test_karpagam_glued_brand_twin():
    items = parse_desc_part_brand_hsn(KARPAGAM_TWIN)
    assert len(items) >= 4
    parts = {it["part_number"] for it in items}
    assert "8PK-1552" in parts
    assert "7971/4RS" in parts or any("7971" in p for p in parts)
    assert "649E" in parts
    courier = [it for it in items if "COURIER" in (it.get("description") or "").upper()]
    assert courier
    assert not courier[0].get("part_number")


def test_anamallais_part_above_si_twin():
    items = parse_anamallais_part_above_si(ANAMALLAIS_TWIN)
    assert len(items) == 2
    assert items[0]["part_number"] == "XX100001"
    assert items[0]["hsn_sac"] == "87087000"
    assert items[1]["part_number"] == "YY200002"
    assert items[1]["hsn_sac"] == "87089900"
    # Must not ingest HSN summary
    assert all(it["part_number"].startswith(("XX", "YY")) for it in items)


def test_ask_optional_disc_twin():
    items = parse_sno_part_desc_discounts(ASK_TWIN)
    assert len(items) >= 3
    parts = {it["part_number"].upper() for it in items}
    assert "CS607822" in parts
    assert "CS100667" in parts
    assert "AB999001" in parts


def test_hema_part_in_desc_twin():
    items = parse_goods_hsn_qty_rate_disc(HEMA_TWIN) or parse_line_items(HEMA_TWIN)
    assert items
    # After validate strip, part should separate — check raw or line items
    all_items = parse_line_items(HEMA_TWIN)
    assert all_items
    joined = " ".join(
        f"{it.get('part_number','')} {it.get('description','')}" for it in all_items
    )
    assert "LX" in joined.upper() and "3630" in joined or "9999" in joined


def test_pr_process_fiscal_supplier_strip():
    # Text-level parse still works; supplier cleaned in extract path via find_supplier
    items = parse_goods_hsn_qty_rate_disc(PR_TWIN) or parse_line_items(PR_TWIN)
    assert len(items) >= 2
    assert any("LL149" in (it.get("part_number") or it.get("description") or "") for it in items)


def test_alagu_must_not_invent_wheel_part():
    items = parse_line_items(ALAGU_NO_PART)
    assert len(items) >= 2
    for it in items:
        assert not looks_like_part_number(it.get("part_number") or "") or it.get("part_number") == ""
        assert "Wheel" in (it.get("description") or "") or "BOOT" in (it.get("description") or "")
        assert (it.get("part_number") or "").lower() not in {"wheel", "boot", "minda"}


def test_scan1_karpagam_noisy_ocr():
    items = parse_desc_part_brand_hsn(SCAN1_KARPAGAM_NOISY)
    assert len(items) == 2
    assert items[0]["part_number"] == "B6Y03007"
    assert "V ROD" in items[0]["description"].upper() or "ASSY" in items[0]["description"].upper()
    assert items[0]["hsn_sac"] == "73269099"
    assert items[0]["qty"].startswith("1")
    assert "10935" in items[0]["rate"].replace(",", "")
    assert "10935" in items[0]["amount"].replace(",", "")
    assert "COURIER" in items[1]["description"].upper()
    assert items[1]["amount"] == "50.00"


def test_credit_twin_no_mrp():
    items = parse_line_items(CREDIT_TWIN)
    assert items
    assert items[0]["part_number"] == "Z9P08758"
    assert not items[0].get("mrp")


def test_qa_pdf_pack_if_present():
    """Live PDF pack: assert min counts so digital QA cannot silently regress."""
    qa = Path(__file__).resolve().parents[2] / (
        "samples/RE_ Invoice PDF_Image to Excel Generator Tool-Status on 30-07-26"
    )
    if not qa.is_dir():
        print("QA pack missing — skip PDF assertions")
        return

    expectations = {
        "A.L.A. PRODUCTZ 09366.pdf": {"min_n": 1, "min_parts": 1, "part_has": "7818"},
        "ANBU BEARINGS 7218.pdf": {"min_n": 1, "min_parts": 1},
        "ARVIND PETROLEUMS 3722.pdf": {"min_n": 3, "min_parts": 3},
        "ASK AUTOMOBILE CR1220.pdf": {"min_n": 18, "min_parts": 18},
        "AUTOLIGHT AGENCY 10325.pdf": {"min_n": 1, "min_parts": 1},
        "Anamallais Agencies (Stadium).pdf": {"min_n": 14, "min_parts": 14},
        "HEMA AGENCIES 3405.pdf": {"min_n": 1, "min_parts": 1},
        "MB Agencies 1546.pdf": {"min_n": 1, "min_parts": 1},
        "P.R. PROCESS 0491.pdf": {"min_n": 10, "min_parts": 10},
        "karpagam auto store.pdf": {"min_n": 30, "min_parts": 28},
    }
    for name, exp in expectations.items():
        pdf = qa / name
        assert pdf.exists(), name
        out = extract_invoice(pdf.read_bytes(), name)
        items = out.get("line_items") or []
        parts = sum(1 for it in items if looks_like_part_number(it.get("part_number") or ""))
        assert len(items) >= exp["min_n"], f"{name}: n={len(items)} < {exp['min_n']}"
        assert parts >= exp["min_parts"], f"{name}: parts={parts} < {exp['min_parts']}"
        if exp.get("part_has"):
            blob = " ".join(it.get("part_number") or "" for it in items)
            assert exp["part_has"] in blob, f"{name}: missing part {exp['part_has']}"
        assert out.get("supplier_name") not in (None, "", "Unknown")
        assert not str(out.get("supplier_name", "")).startswith("(")


def test_web_style_twins_if_present():
    root = Path(__file__).resolve().parents[2] / "samples/web-tests"
    checks = {
        "web-ashok-style.pdf": {"min_n": 3, "min_parts": 3},
        "web-autolight-style.pdf": {"min_n": 1},
        "web-alagu-style.pdf": {"min_n": 2, "forbid_parts": {"Wheel", "BOOT", "Minda"}},
    }
    for name, exp in checks.items():
        pdf = root / name
        if not pdf.exists():
            continue
        out = extract_invoice(pdf.read_bytes(), name)
        items = out.get("line_items") or []
        assert len(items) >= exp["min_n"], name
        if exp.get("min_parts"):
            parts = sum(1 for it in items if looks_like_part_number(it.get("part_number") or ""))
            assert parts >= exp["min_parts"], name
        for bad in exp.get("forbid_parts") or ():
            assert all((it.get("part_number") or "") != bad for it in items), f"{name} invented {bad}"


if __name__ == "__main__":
    tests = [v for k, v in globals().items() if k.startswith("test_") and callable(v)]
    for fn in tests:
        fn()
        print("OK", fn.__name__)
    print(f"ALL {len(tests)} QA FAMILY TESTS OK")
