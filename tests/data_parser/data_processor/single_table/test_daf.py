from decimal import Decimal
from typing import Any, Dict, List

from core.data_parser.data_processor.single_table.daf import perform_DAF_compounding


def test_daf_empty_input():
    """Empty list returns single default empty quantized DAF row."""
    res = perform_DAF_compounding([])
    assert len(res) == 1
    assert res[0] == {
        "col_po": "",
        "col_item": "",
        "col_desc": "",
        "col_qty_sf": Decimal(0),
        "col_amount": Decimal(0),
        "col_net": Decimal(0),
        "col_dc": "",
        "material_type": "normal",
    }


def test_daf_material_split_buffalo_and_normal():
    """Has descriptions: groups by material with Buffalo first, normal second."""
    data: List[Dict[str, Any]] = [
        {
            "col_po": "PO-101",
            "col_item": "ITEM-A",
            "col_desc": "BUFFALO EMBOSSED LEATHER",
            "col_qty_sf": Decimal("100.5"),
            "col_amount": Decimal("1005.0"),
            "col_net": Decimal("200.0"),
            "col_dc": "DC-NORTH",
        },
        {
            "col_po": "PO-102",
            "col_item": "ITEM-B",
            "col_desc": "COW SPLIT LEATHER",
            "col_qty_sf": Decimal("50.0"),
            "col_amount": Decimal("500.0"),
            "col_net": Decimal("100.0"),
            "col_dc": "DC-SOUTH",
        },
        {
            "col_po": "PO-103",
            "col_item": "ITEM-C",
            "col_desc": "DARK BUFFALO LEATHER",
            "col_qty_sf": Decimal("40.0"),
            "col_amount": Decimal("400.0"),
            "col_net": Decimal("80.0"),
            "col_dc": "DC-NORTH",
        },
    ]

    res = perform_DAF_compounding(data)
    assert len(res) == 2

    # Buffalo group is first
    assert res[0]["material_type"] == "buffalo"
    assert res[0]["col_qty_sf"] == Decimal("140.5")
    assert res[0]["col_amount"] == Decimal("1405.0")
    assert res[0]["col_net"] == Decimal("280.0")
    assert res[0]["col_po"] == "PO-101/PO-103"
    assert res[0]["col_item"] == "ITEM-A/ITEM-C"
    assert res[0]["col_desc"] == "BUFFALO EMBOSSED LEATHER\nDARK BUFFALO LEATHER"
    assert res[0]["col_dc"] == "DC-NORTH"

    # Normal group is second
    assert res[1]["material_type"] == "normal"
    assert res[1]["col_qty_sf"] == Decimal("50.0")
    assert res[1]["col_amount"] == Decimal("500.0")
    assert res[1]["col_net"] == Decimal("100.0")
    assert res[1]["col_po"] == "PO-102"
    assert res[1]["col_item"] == "ITEM-B"
    assert res[1]["col_desc"] == "COW SPLIT LEATHER"
    assert res[1]["col_dc"] == "DC-SOUTH"


def test_daf_po_chunk_under_7_pos():
    """No descriptions and <= 7 POs: outputs single po_chunk row with exactly 7 POs."""
    data: List[Dict[str, Any]] = [
        {
            "col_po": f"PO-{i}",
            "col_item": f"ITEM-{i}",
            "col_desc": "",
            "col_qty_sf": Decimal("10.0"),
            "col_amount": Decimal("100.0"),
            "col_net": Decimal("20.0"),
            "col_dc": "DC-1",
        }
        for i in range(1, 8)  # Exactly 7 POs <= 7
    ]

    res = perform_DAF_compounding(data)
    assert len(res) == 1
    assert res[0]["material_type"] == "po_chunk"
    assert res[0]["col_po"] == "PO-1/PO-2\nPO-3/PO-4\nPO-5/PO-6\nPO-7"
    assert res[0]["col_item"] == "ITEM-1/ITEM-2\nITEM-3/ITEM-4\nITEM-5/ITEM-6\nITEM-7"
    assert res[0]["col_desc"] == ""
    assert res[0]["col_qty_sf"] == Decimal("70.0")
    assert res[0]["col_amount"] == Decimal("700.0")
    assert res[0]["col_net"] == Decimal("140.0")


def test_daf_po_chunk_over_7_pos():
    """No descriptions and > 7 POs: splits into two po_chunk rows (8 POs even 4/4 split, 9 POs odd 5/4 split)."""
    # 8 POs: exact boundary cutoff for >7, 4/4 even split
    data_8: List[Dict[str, Any]] = [
        {
            "col_po": f"PO-{i:02d}",
            "col_item": f"ITEM-{i:02d}",
            "col_desc": None,
            "col_qty_sf": Decimal("10.0"),
            "col_amount": Decimal("100.0"),
            "col_net": Decimal("20.0"),
            "col_dc": "DC-1",
        }
        for i in range(1, 9)  # 8 POs > 7
    ]

    res_8 = perform_DAF_compounding(data_8)
    assert len(res_8) == 2
    assert res_8[0]["material_type"] == "po_chunk"
    assert res_8[1]["material_type"] == "po_chunk"
    assert res_8[0]["col_po"] == "PO-01/PO-02\nPO-03/PO-04"
    assert res_8[0]["col_qty_sf"] == Decimal("40.0")
    assert res_8[0]["col_amount"] == Decimal("400.0")
    assert res_8[0]["col_net"] == Decimal("80.0")
    assert res_8[1]["col_po"] == "PO-05/PO-06\nPO-07/PO-08"
    assert res_8[1]["col_qty_sf"] == Decimal("40.0")
    assert res_8[1]["col_amount"] == Decimal("400.0")
    assert res_8[1]["col_net"] == Decimal("80.0")
    assert res_8[0]["col_qty_sf"] + res_8[1]["col_qty_sf"] == Decimal("80.0")
    assert res_8[0]["col_amount"] + res_8[1]["col_amount"] == Decimal("800.0")

    # 9 POs: odd split (5/4)
    data_9: List[Dict[str, Any]] = [
        {
            "col_po": f"PO-{i:02d}",
            "col_item": f"ITEM-{i:02d}",
            "col_desc": None,
            "col_qty_sf": Decimal("10.0"),
            "col_amount": Decimal("100.0"),
            "col_net": Decimal("20.0"),
            "col_dc": "DC-1",
        }
        for i in range(1, 10)  # 9 POs > 7
    ]

    res = perform_DAF_compounding(data_9)
    assert len(res) == 2
    assert res[0]["material_type"] == "po_chunk"
    assert res[1]["material_type"] == "po_chunk"

    # Chunk 0: 5 POs ((9 + 1) // 2 = 5)
    assert res[0]["col_po"] == "PO-01/PO-02\nPO-03/PO-04\nPO-05"
    assert res[0]["col_qty_sf"] == Decimal("50.0")
    assert res[0]["col_amount"] == Decimal("500.0")
    assert res[0]["col_net"] == Decimal("100.0")

    # Chunk 1: 4 POs
    assert res[1]["col_po"] == "PO-06/PO-07\nPO-08/PO-09"
    assert res[1]["col_qty_sf"] == Decimal("40.0")
    assert res[1]["col_amount"] == Decimal("400.0")
    assert res[1]["col_net"] == Decimal("80.0")

    # Total sum across both chunks equals total input
    total_qty = res[0]["col_qty_sf"] + res[1]["col_qty_sf"]
    total_amt = res[0]["col_amount"] + res[1]["col_amount"]
    assert total_qty == Decimal("90.0")
    assert total_amt == Decimal("900.0")


def test_daf_formatting_and_separators():
    """Tests chunk formatting, custom separators, and fallback when chunk_size <= 0."""
    data: List[Dict[str, Any]] = [
        {
            "col_po": f"PO-{i}",
            "col_item": f"ITEM-{i}",
            "col_desc": "LEATHER",
            "col_qty_sf": Decimal("10.0"),
            "col_amount": Decimal("100.0"),
            "col_net": Decimal("20.0"),
        }
        for i in range(1, 5)
    ]

    # Custom chunk_size and separators
    res_custom = perform_DAF_compounding(
        data,
        chunk_size=2,
        item_separator=" ; ",
        line_separator=" | "
    )
    assert len(res_custom) == 1
    assert res_custom[0]["col_po"] == "PO-1 ; PO-2 | PO-3 ; PO-4"

    # Invalid chunk_size <= 0 defaults to 2
    res_neg = perform_DAF_compounding(data, chunk_size=-1)
    assert len(res_neg) == 1
    assert res_neg[0]["col_po"] == "PO-1/PO-2\nPO-3/PO-4"

    res_zero = perform_DAF_compounding(data, chunk_size=0)
    assert len(res_zero) == 1
    assert res_zero[0]["col_po"] == "PO-1/PO-2\nPO-3/PO-4"


def test_daf_decimals_quantization():
    """Verify decimals_map quantizes Decimal results properly."""
    data: List[Dict[str, Any]] = [
        {
            "col_po": "PO-101",
            "col_item": "ITEM-A",
            "col_desc": "BUFFALO LEATHER",
            "col_qty_sf": Decimal("100.5555"),
            "col_amount": Decimal("1005.149"),
            "col_net": Decimal("200.846"),
            "col_dc": "DC-NORTH",
        },
        {
            "col_po": "PO-102",
            "col_item": "ITEM-B",
            "col_desc": "BUFFALO LEATHER",
            "col_qty_sf": Decimal("50.1111"),
            "col_amount": Decimal("500.251"),
            "col_net": Decimal("100.101"),
            "col_dc": "DC-NORTH",
        },
    ]

    decimals_map = {
        "col_qty_sf": 3,
        "col_amount": 1,
        # col_net omitted to verify default 2 decimal places
    }

    res = perform_DAF_compounding(data, decimals_map=decimals_map)
    assert len(res) == 1
    assert res[0]["col_qty_sf"] == Decimal("150.667")
    assert res[0]["col_amount"] == Decimal("1505.4")
    assert res[0]["col_net"] == Decimal("300.95")
