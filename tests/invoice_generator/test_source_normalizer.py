import pytest
from typing import Any, Dict, List

from core.invoice_generator.mappers import InvoiceMode
from core.invoice_generator.mappers.rules import (
    get_fallback_value,
    resolve_mode_formula,
)
from core.invoice_generator.mappers.source_normalizer import (
    normalize_source_to_rows,
    extract_cell_value,
)
from core.invoice_generator.mappers.transforms import prepare_data_rows


# ==============================================================================
# 1. Tests for normalize_source_to_rows
# ==============================================================================


def test_normalize_list_of_dicts():
    """Verify list of dicts returns shallow copies of records, filtering non-dict elements."""
    record_a = {"col_item": "Item A", "col_qty": 10}
    record_b = {"col_item": "Item B", "col_qty": 20}
    input_records = [record_a, "invalid_element", None, record_b]

    rows = normalize_source_to_rows(input_records)

    assert len(rows) == 2
    assert rows[0] == {"col_item": "Item A", "col_qty": 10}
    assert rows[1] == {"col_item": "Item B", "col_qty": 20}

    # Verify shallow copy
    assert rows[0] is not record_a
    assert rows[1] is not record_b


def test_normalize_scalar_dict():
    """Verify scalar dict (no lists) returns a single-row list."""
    scalar_data = {
        "col_item": "Single Item",
        "col_qty": 50,
        "col_price": 9.99,
    }
    rows = normalize_source_to_rows(scalar_data)

    assert len(rows) == 1
    assert rows[0] == scalar_data
    # Verify copy
    assert rows[0] is not scalar_data


@pytest.mark.parametrize(
    "invalid_input",
    [
        {},
        [],
        None,
        "",
        "some_string",
        12345,
        0,
        False,
        True,
        object(),
    ],
)
def test_normalize_empty_and_invalid_types(invalid_input: Any):
    """Verify empty dict, empty list, None, and invalid types return empty list."""
    rows = normalize_source_to_rows(invalid_input)
    assert rows == []


# ==============================================================================
# 2. Tests for extract_cell_value
# ==============================================================================


def test_extract_cell_value_rule_column():
    """Verify value extraction when rule specifies 'column'."""
    row = {"raw_col_name": "Column Value"}
    rule = {"column": "raw_col_name"}
    assert extract_cell_value(row, "target_key", rule) == "Column Value"


def test_extract_cell_value_column_precedence_over_rule_key():
    """Verify 'column' mapping takes precedence over fallback rule_key."""
    row = {
        "col_mapped": "from column",
        "target_key": "from fallback",
    }
    rule = {"column": "col_mapped"}
    assert extract_cell_value(row, "target_key", rule) == "from column"


def test_extract_cell_value_fallback_rule_key():
    """Verify fallback to rule_key when rule mappings are absent or rule is None/empty."""
    row = {"fallback_key": "Fallback Value"}

    # Rule is None
    assert extract_cell_value(row, "fallback_key", None) == "Fallback Value"

    # Rule is empty dict
    assert extract_cell_value(row, "fallback_key", {}) == "Fallback Value"

    # Rule has non-existent mapped key -> falls back to rule_key
    assert extract_cell_value(row, "fallback_key", {"column": "non_existent"}) == "Fallback Value"


def test_extract_cell_value_skips_none_and_empty_string():
    """Verify None and empty string values are skipped to try fallback rule_key."""
    # primary mapped key is None -> falls back to rule_key
    row_none = {"col_empty": None, "target_col": "Valid Fallback"}
    rule = {"column": "col_empty"}
    assert extract_cell_value(row_none, "target_col", rule) == "Valid Fallback"

    # primary mapped key is "" -> falls back to rule_key
    row_empty_str = {"col_empty": "", "target_col": "Valid Fallback"}
    assert extract_cell_value(row_empty_str, "target_col", rule) == "Valid Fallback"


def test_extract_cell_value_preserves_falsy_valid_values():
    """Verify numeric 0, 0.0, and boolean False are treated as valid values, not skipped."""
    assert extract_cell_value({"num": 0}, "num", None) == 0
    assert extract_cell_value({"num": 0.0}, "num", None) == 0.0
    assert extract_cell_value({"flag": False}, "flag", None) is False


def test_extract_cell_value_missing_returns_none():
    """Verify None is returned when keys are missing or all candidates contain None/empty string."""
    assert extract_cell_value({"a": "val"}, "b", None) is None
    assert extract_cell_value({"a": None}, "a", None) is None
    assert extract_cell_value({"a": ""}, "a", None) is None
    assert extract_cell_value({"a": None, "b": ""}, "b", {"column": "a"}) is None

    # Invalid row type returns None
    assert extract_cell_value(None, "key", None) is None
    assert extract_cell_value([], "key", None) is None
    assert extract_cell_value("invalid", "key", None) is None


# ==============================================================================
# 3. Integration with prepare_data_rows
# ==============================================================================


def test_prepare_data_rows_integration_list_of_dicts():
    """Verify prepare_data_rows works with list of dicts (records)."""
    records_source = [
        {"raw_name": "Leather Black", "raw_qty": "50"},
        {"raw_name": "Leather Brown", "raw_qty": "75"},
    ]
    mapping_rules = {
        "col_item": {"column": "raw_name"},
        "col_qty_sf": {"column": "raw_qty"},
    }
    static_values = {"col_origin": "Cambodia"}

    rows, count = prepare_data_rows(
        data_source=records_source,
        dynamic_mapping_rules=mapping_rules,
        static_value_map=static_values,
    )

    assert count == 2
    assert len(rows) == 2

    assert rows[0]["col_item"] == "Leather Black"
    assert rows[0]["col_qty_sf"] == "50"
    assert rows[0]["col_origin"] == "Cambodia"

    assert rows[1]["col_item"] == "Leather Brown"
    assert rows[1]["col_qty_sf"] == "75"
    assert rows[1]["col_origin"] == "Cambodia"


def test_prepare_data_rows_integration_scalar_dict():
    """Verify prepare_data_rows works with a scalar dict input producing a single row."""
    scalar_source = {
        "col_item": "Single Special Item",
        "col_qty_sf": "1000",
        "col_amount": "5000.00",
    }
    mapping_rules = {
        "col_item": {"column": "col_item"},
        "col_qty_sf": {"column": "col_qty_sf"},
        "col_amount": {"column": "col_amount"},
    }

    rows, count = prepare_data_rows(
        data_source=scalar_source,
        dynamic_mapping_rules=mapping_rules,
    )

    assert count == 1
    assert len(rows) == 1
    assert rows[0]["col_item"] == "Single Special Item"
    assert rows[0]["col_qty_sf"] == "1000"
    assert rows[0]["col_amount"] == "5000.00"


def test_prepare_data_rows_integration_empty_and_none():
    """Verify prepare_data_rows handles empty dict and None gracefully."""
    mapping_rules = {"col_item": {"column": "col_item"}}

    rows_empty, count_empty = prepare_data_rows(
        data_source={},
        dynamic_mapping_rules=mapping_rules,
    )
    assert count_empty == 0
    assert rows_empty == []

    rows_none, count_none = prepare_data_rows(
        data_source=None,
        dynamic_mapping_rules=mapping_rules,
    )
    assert count_none == 0
    assert rows_none == []


def test_prepare_data_rows_preserves_falsey_cell_values():
    """Verify numeric 0, 0.0, and boolean False are preserved and not dropped."""
    source_records = [
        {"col_qty": 0, "col_active": False, "col_price": 0.0},
    ]
    mapping_rules = {
        "col_qty": {"column": "col_qty"},
        "col_active": {"column": "col_active"},
        "col_price": {"column": "col_price"},
    }
    rows, count = prepare_data_rows(
        data_source=source_records,
        dynamic_mapping_rules=mapping_rules,
    )
    assert count == 1
    assert rows[0]["col_qty"] == 0
    assert rows[0]["col_active"] is False
    assert rows[0]["col_price"] == 0.0


def test_prepare_data_rows_col_desc_fallback_with_mapped_col_po():
    """Verify col_desc fallback behaves correctly with mapped col_po and skips when col_po is None or empty."""
    source = [
        {"raw_po": "PO-999", "col_desc": None},       # valid mapped PO -> fallback applies
        {"raw_po": "", "col_desc": None},             # empty mapped PO -> fallback skipped
        {"raw_po": None, "col_desc": None},           # None mapped PO -> fallback skipped
        {"raw_po": "PO-888", "col_desc": "Given"},    # explicit desc present -> not replaced
    ]
    mapping_rules = {
        "col_desc": {"column": "col_desc", "fallback": "Fallback Description"},
        "col_po": {"column": "raw_po"},
    }
    rows, count = prepare_data_rows(
        data_source=source,
        dynamic_mapping_rules=mapping_rules,
    )
    assert count == 4
    # Row 0: fallback applied because col_po is present via mapped raw_po
    assert rows[0]["col_po"] == "PO-999"
    assert rows[0]["col_desc"] == "Fallback Description"

    # Row 1: empty PO -> col_desc skipped
    assert "col_desc" not in rows[1]

    # Row 2: None PO -> col_desc skipped
    assert "col_desc" not in rows[2]

    # Row 3: explicit desc preserved
    assert rows[3]["col_po"] == "PO-888"
    assert rows[3]["col_desc"] == "Given"


def test_invoice_mode_initialization():
    """Verify clean InvoiceMode initialization via explicit format, pricing_basis, and from_dict."""
    # Direct construction with defaults
    mode_std = InvoiceMode()
    assert mode_std.invoice_format == "standard"
    assert mode_std.pricing_basis == "sqft"
    assert mode_std.DAF_mode is False
    assert mode_std.custom_mode is False
    assert mode_std.item_merged_mode is False
    assert mode_std.pricing_net_weight is False

    # Direct construction with explicit invoice_format and pricing_basis
    mode_daf = InvoiceMode(invoice_format="daf", pricing_basis="net_weight")
    assert mode_daf.invoice_format == "daf"
    assert mode_daf.pricing_basis == "net_weight"
    assert mode_daf.DAF_mode is True
    assert mode_daf.pricing_net_weight is True
    assert mode_daf.custom_mode is False
    assert mode_daf.item_merged_mode is False

    mode_custom = InvoiceMode(invoice_format="custom")
    assert mode_custom.invoice_format == "custom"
    assert mode_custom.pricing_basis == "sqft"
    assert mode_custom.custom_mode is True
    assert mode_custom.DAF_mode is False
    assert mode_custom.item_merged_mode is False

    mode_merged = InvoiceMode(invoice_format="item_merged")
    assert mode_merged.invoice_format == "item_merged"
    assert mode_merged.pricing_basis == "sqft"
    assert mode_merged.item_merged_mode is True
    assert mode_merged.custom_mode is False
    assert mode_merged.DAF_mode is False

    # from_dict with explicit invoice_format or format
    assert InvoiceMode.from_dict({"invoice_format": "daf"}).DAF_mode is True
    assert InvoiceMode.from_dict({"invoice_format": "custom"}).custom_mode is True
    assert InvoiceMode.from_dict({"invoice_format": "item_merged"}).item_merged_mode is True
    assert InvoiceMode.from_dict({"invoice_format": "standard"}).invoice_format == "standard"
    assert InvoiceMode.from_dict({"format": "daf"}).DAF_mode is True
    assert InvoiceMode.from_dict({"format": "custom"}).custom_mode is True
    assert InvoiceMode.from_dict({"format": "item_merged"}).item_merged_mode is True

    # from_dict with pricing_basis and pricing_net_weight
    mode_dict_daf = InvoiceMode.from_dict({"invoice_format": "daf", "pricing_basis": "net_weight"})
    assert mode_dict_daf.DAF_mode is True
    assert mode_dict_daf.pricing_net_weight is True

    assert InvoiceMode.from_dict({"pricing_net_weight": True}).pricing_net_weight is True

    # from_dict idempotence and empty defaults
    assert InvoiceMode.from_dict(mode_daf) is mode_daf
    assert InvoiceMode.from_dict(None).invoice_format == "standard"
    assert InvoiceMode.from_dict({}).invoice_format == "standard"


def test_invoice_mode_invalid_format():
    """Verify ValueError is raised on invalid invoice_format and pricing_basis."""
    with pytest.raises(ValueError, match="Invalid invoice_format"):
        InvoiceMode(invoice_format="invalid")

    with pytest.raises(ValueError, match="Invalid pricing_basis"):
        InvoiceMode(pricing_basis="invalid")

    with pytest.raises(ValueError, match="Invalid invoice_format"):
        InvoiceMode.from_dict({"invoice_format": "unknown"})

    with pytest.raises(ValueError, match="Invalid pricing_basis"):
        InvoiceMode.from_dict({"pricing_basis": "unknown"})

    with pytest.raises(ValueError, match="Invalid invoice_format"):
        InvoiceMode.from_dict({"format": "unknown"})


def test_invoice_mode_uppercase_normalization():
    """Verify uppercase format strings 'DAF', 'CUSTOM' normalize correctly."""
    mode_daf = InvoiceMode(invoice_format="DAF", pricing_basis="NET_WEIGHT")
    assert mode_daf.invoice_format == "daf"
    assert mode_daf.pricing_basis == "net_weight"
    assert mode_daf.DAF_mode is True
    assert mode_daf.pricing_net_weight is True

    mode_custom = InvoiceMode.from_dict({"invoice_format": "CUSTOM"})
    assert mode_custom.invoice_format == "custom"
    assert mode_custom.custom_mode is True

    mode_dict_format = InvoiceMode.from_dict({"format": "DAF"})
    assert mode_dict_format.invoice_format == "daf"
    assert mode_dict_format.DAF_mode is True


def test_invoice_mode_from_dict_non_dict():
    """Verify ValueError is raised when non-dict is passed to from_dict."""
    with pytest.raises(ValueError, match="Expected dict or InvoiceMode"):
        InvoiceMode.from_dict("invalid")

    with pytest.raises(ValueError, match="Expected dict or InvoiceMode"):
        InvoiceMode.from_dict(123)

    with pytest.raises(ValueError, match="Expected dict or InvoiceMode"):
        InvoiceMode.from_dict([1, 2, 3])

    for falsy_val in [0, False, [], ""]:
        with pytest.raises(ValueError, match="Expected dict or InvoiceMode"):
            InvoiceMode.from_dict(falsy_val)


def test_resolve_mode_formula_and_get_fallback_value():
    """Verify resolve_mode_formula and get_fallback_value mode resolution and validation."""
    rule = {
        "column": "col_test",
        "formula": {
            "standard": "{col_a} + {col_b}",
            "custom": "{col_a} * {col_b}",
            "daf": "{col_a} - {col_b}",
            # item_merged is omitted to verify it does not cross-fallback to custom
        },
        "fallback": {
            "standard": "Fallback Standard",
            "custom": "Fallback Custom",
            "daf": "Fallback DAF",
        },
    }

    # resolve_mode_formula for item_merged does NOT cross-fallback to custom
    assert resolve_mode_formula(rule, invoice_format="item_merged") is None
    assert resolve_mode_formula(rule, invoice_format="custom") == "{col_a} * {col_b}"
    assert resolve_mode_formula(rule, invoice_format="daf") == "{col_a} - {col_b}"
    assert resolve_mode_formula(rule, invoice_format="standard") == "{col_a} + {col_b}"

    # Uppercase normalization
    assert resolve_mode_formula(rule, invoice_format="DAF") == "{col_a} - {col_b}"
    assert resolve_mode_formula(rule, invoice_format="CUSTOM") == "{col_a} * {col_b}"

    # get_fallback_value returns None for unsupported nested dict fallback
    assert get_fallback_value(rule, invoice_format="standard") is None
    assert get_fallback_value(rule, invoice_format="custom") is None
    assert get_fallback_value(rule, invoice_format="daf") is None
    assert get_fallback_value(rule, invoice_format="item_merged") is None
    assert get_fallback_value(rule, invoice_format="DAF") is None
    assert get_fallback_value(rule, invoice_format="CUSTOM") is None

    # Flat string fallback returns same value across all modes
    flat_rule = {"fallback": "COW LEATHER"}
    for fmt in ("standard", "custom", "daf", "item_merged"):
        assert get_fallback_value(flat_rule, invoice_format=fmt) == "COW LEATHER"

    # Legacy flat keys are unsupported and return None
    legacy_rule = {"fallback_on_none": "LEGACY LEATHER"}
    assert get_fallback_value(legacy_rule, invoice_format="standard") is None
    assert get_fallback_value(legacy_rule, invoice_format="custom") is None

    legacy_daf_rule = {"fallback_on_DAF": "LEGACY DAF"}
    assert get_fallback_value(legacy_daf_rule, invoice_format="daf") is None

    legacy_custom_rule = {"fallback_on_custom": "LEGACY CUSTOM"}
    assert get_fallback_value(legacy_custom_rule, invoice_format="custom") is None

    # Invalid invoice_format raises ValueError even if rule is None or empty dict
    with pytest.raises(ValueError, match="Invalid invoice_format"):
        resolve_mode_formula(rule, invoice_format="invalid")

    with pytest.raises(ValueError, match="Invalid invoice_format"):
        resolve_mode_formula(None, invoice_format="invalid")

    with pytest.raises(ValueError, match="Invalid invoice_format"):
        resolve_mode_formula({}, invoice_format="invalid")


test_mode_rules_resolution_and_validation = test_resolve_mode_formula_and_get_fallback_value



def test_prepare_data_rows_with_invoice_mode():
    """Verify passing InvoiceMode directly or as a dict to prepare_data_rows mode parameter."""
    source = [{"raw_po": "PO-1", "col_desc": None}]
    mapping_rules = {
        "col_po": {"column": "raw_po"},
        "col_desc": {
            "column": "col_desc",
            "fallback": "Universal Desc",
        },
        "col_total": {
            "formula": "{col_qty_sf} * {col_price}",
        },
    }

    # Passing InvoiceMode directly
    rows_direct, count_direct = prepare_data_rows(
        source,
        mapping_rules,
        mode=InvoiceMode(invoice_format="daf"),
    )
    assert count_direct == 1
    assert rows_direct[0]["col_desc"] == "Universal Desc"

    # Passing dict to mode={'invoice_format': 'daf'}
    rows_dict, count_dict = prepare_data_rows(
        source,
        mapping_rules,
        mode={"invoice_format": "daf"},
    )
    assert count_dict == 1
    assert rows_dict[0]["col_desc"] == "Universal Desc"

    # Passing string mode='daf'
    rows_str, count_str = prepare_data_rows(
        source,
        mapping_rules,
        mode="daf",
    )
    assert count_str == 1
    assert rows_str[0]["col_desc"] == "Universal Desc"

    # custom mode direct, dict, and string
    rows_custom_direct, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode=InvoiceMode(invoice_format="custom"),
    )
    assert rows_custom_direct[0]["col_desc"] == "Universal Desc"

    rows_custom_dict, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode={"invoice_format": "custom"},
    )
    assert rows_custom_dict[0]["col_desc"] == "Universal Desc"

    rows_custom_str, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode="custom",
    )
    assert rows_custom_str[0]["col_desc"] == "Universal Desc"

    # item_merged mode direct, dict, and string
    rows_merged_direct, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode=InvoiceMode(invoice_format="item_merged"),
    )
    assert rows_merged_direct[0]["col_desc"] == "Universal Desc"

    rows_merged_dict, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode={"invoice_format": "item_merged"},
    )
    assert rows_merged_dict[0]["col_desc"] == "Universal Desc"

    rows_merged_str, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode="item_merged",
    )
    assert rows_merged_str[0]["col_desc"] == "Universal Desc"

    # pricing_net_weight direct and dict (replaces {col_qty_sf} with {col_net})
    rows_pricing_direct, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode=InvoiceMode(pricing_basis="net_weight"),
    )
    assert rows_pricing_direct[0]["col_total"]["inputs"] == ["col_net", "col_price"]

    rows_pricing_dict, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode={"pricing_basis": "net_weight"},
    )
    assert rows_pricing_dict[0]["col_total"]["inputs"] == ["col_net", "col_price"]

    rows_pricing_compat, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode={"pricing_net_weight": True},
    )
    assert rows_pricing_compat[0]["col_total"]["inputs"] == ["col_net", "col_price"]

    rows_pricing_false, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode={"pricing_net_weight": False},
    )
    assert rows_pricing_false[0]["col_total"]["inputs"] == ["col_qty_sf", "col_price"]


def test_prepare_data_rows_pricing_net_weight_dict_formula():
    """Verify pricing_net_weight updates both template and inputs list for dict formulas."""
    source = [{"col_qty_sf": 100, "col_net": 80.5, "col_price": 12.0}]
    mapping_rules = {
        "col_total": {
            "formula": {
                "standard": {
                    "template": "{col_qty_sf} * {col_price}",
                    "inputs": ["col_qty_sf", "col_price"],
                }
            }
        },
        "col_total_ref": {
            "formula": {
                "standard": {
                    "template": "{col_ref_0} * {col_ref_1}",
                    "inputs": ["col_qty_sf", "col_price"],
                }
            }
        },
    }
    # When pricing_net_weight is True, template and inputs should replace col_qty_sf with col_net
    rows_enabled, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode=InvoiceMode(pricing_basis="net_weight"),
    )
    assert rows_enabled[0]["col_total"]["template"] == "{col_net} * {col_price}"
    assert rows_enabled[0]["col_total"]["inputs"] == ["col_net", "col_price"]
    assert rows_enabled[0]["col_total_ref"]["template"] == "{col_ref_0} * {col_ref_1}"
    assert rows_enabled[0]["col_total_ref"]["inputs"] == ["col_net", "col_price"]

    # When pricing_net_weight is False, original formula remains intact
    rows_disabled, _ = prepare_data_rows(
        source,
        mapping_rules,
        mode=InvoiceMode(pricing_basis="sqft"),
    )
    assert rows_disabled[0]["col_total"]["template"] == "{col_qty_sf} * {col_price}"
    assert rows_disabled[0]["col_total"]["inputs"] == ["col_qty_sf", "col_price"]
    assert rows_disabled[0]["col_total_ref"]["inputs"] == ["col_qty_sf", "col_price"]



