import pytest
from typing import Dict, Any

from core.invoice_generator.mappers import (
    TableDataMapper,
    ResolvedTableData,
    extract_table_data,
    resolve_static_placeholders,
    populate_static_content,
    extract_summaries,
    format_pallet_counts
)


def test_helpers_extract_table_data():
    """Verify that extract_table_data correctly extracts data and cleans stringified tuple keys."""
    # List data source
    list_source = [{"col_desc": "Buffalo"}]
    assert extract_table_data(list_source, "aggregation") == list_source

    # Stringified tuple keys in dict
    dict_source_stringified = {
        "('BUFFALO', 'COW')": "value1",
        "normal_key": "value2"
    }
    extracted = extract_table_data(dict_source_stringified, "aggregation")
    assert ("BUFFALO", "COW") in extracted
    assert extracted[("BUFFALO", "COW")] == "value1"
    assert extracted["normal_key"] == "value2"

    # None data source
    assert extract_table_data(None, "aggregation") is None


def test_helpers_populate_static_content():
    """Verify resolve_static_placeholders and populate_static_content sequence."""
    data_rows = [{"col_item": "dynamic1"}, {"col_item": "dynamic2"}]
    static_payload = {"col_static": ["Static1", "{col_desc_fallback}"]}

    resolved_static = resolve_static_placeholders(
        static_payload=static_payload,
        desc_fallback_str="Standard Desc"
    )
    assert resolved_static["col_static"][1] == "Standard Desc"

    populate_static_content(
        data_rows=data_rows,
        static_payload=resolved_static
    )

    # Values should be merged into col_static
    assert data_rows[0]["col_static"] == "Static1"
    assert data_rows[1]["col_static"] == "Standard Desc"

    # Test extending rows if we have more static values than data rows
    data_rows_short = [{"col_item": "dynamic1"}]
    static_payload_long = {"col_static": ["Static1", "Static2"]}
    populate_static_content(
        data_rows=data_rows_short,
        static_payload=static_payload_long
    )
    assert len(data_rows_short) == 2
    assert data_rows_short[0]["col_static"] == "Static1"
    assert data_rows_short[1]["col_static"] == "Static2"


def test_helpers_extract_summaries():
    """Verify that extract_summaries handles various source types and footer data formats."""
    # Dict source legacy path
    data_source_dict = {
        "leather_summary": "l_sum",
        "weight_summary": "w_sum",
        "pallet_summary_total": 4
    }
    l, w, p = extract_summaries(data_source_dict, footer_data={})
    assert l == "l_sum"
    assert w == "w_sum"
    assert p == 4

    # Pre-calculated in footer_data grand_total
    footer_data = {
        "grand_total": {
            "col_pallet_count": 10
        }
    }
    _, _, p = extract_summaries(None, footer_data)
    assert p == 10

    # Multi-table table_totals list
    footer_data_multi = {
        "table_totals": [
            {"col_pallet_count": 3},
            {"col_pallet_count": 5}
        ]
    }
    _, _, p0 = extract_summaries(None, footer_data_multi)
    assert p0 == 3


def test_helpers_format_pallet_counts():
    """Verify that format_pallet_counts carries values forward for vertical merging."""
    data_rows = [
        {"col_item": "Item 1", "col_pallet_count": "1-5"},
        {"col_item": "Item 2", "col_pallet_count": 0},
        {"col_item": "Item 3", "col_pallet_count": "2-5"}
    ]
    format_pallet_counts(
        data_rows=data_rows,
        num_data_rows=3,
        pallet_col_id="col_pallet_count"
    )

    assert data_rows[0]["col_pallet_count"] == "1-5"
    assert data_rows[1]["col_pallet_count"] == "1-5"
    assert data_rows[2]["col_pallet_count"] == "2-5"


def test_mapping_context_and_rule_engine():
    """Verify MappingContext initialization and RuleEngine evaluation."""
    from core.invoice_generator.mappers import MappingContext, RuleEngine

    context = MappingContext(
        data_source_type="aggregation",
        data_source=[{"col_a": "val1"}],
        DAF_mode=True,
        custom_mode=False
    )
    assert context.DAF_mode is True
    assert context.custom_mode is False

    row_dict = {}
    rule = {
        "fallback": {
            "daf": "DAF Value",
            "standard": "Standard Value"
        }
    }
    RuleEngine.evaluate_column_rule("col_test", rule, row_dict, context)
    assert row_dict["col_test"] == "DAF Value"


def test_table_binding_single_input_mapping():
    """Verify single-input TableBinding construction and mapper resolution."""
    from core.invoice_generator.mappers import TableBinding, TableDataMapper

    data_config = {
        "data_source_type": "aggregation",
        "data_source": [{"col_item": "leather_bag"}],
        "mapping_rules": {"col_item": {"field": "col_item"}}
    }
    context_config = {"DAF_mode": False}

    binding = TableBinding.from_bundles(data_config, context_config)
    assert binding.data == [{"col_item": "leather_bag"}]

    mapper = TableDataMapper(binding=binding)
    assert mapper.binding is binding
    assert mapper.data_source == binding.data


def test_populate_static_content_empty_data_rows():
    """Verify resolve_static_placeholders + populate_static_content when data_rows is initially empty."""
    data_rows = []
    static_payload = {"col_static": ["VENDOR#:", "Des: {col_desc_fallback}", "MADE IN CAMBODIA"]}

    resolved_static = resolve_static_placeholders(
        static_payload=static_payload,
        desc_fallback_str="Standard Cow Leather"
    )

    populate_static_content(
        data_rows=data_rows,
        static_payload=resolved_static
    )

    assert len(data_rows) == 3
    assert data_rows[0]["col_static"] == "VENDOR#:"
    assert data_rows[1]["col_static"] == "Des: Standard Cow Leather"
    assert data_rows[2]["col_static"] == "MADE IN CAMBODIA"



