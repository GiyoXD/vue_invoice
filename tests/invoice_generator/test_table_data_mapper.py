import pytest
from typing import Dict, Any

from core.invoice_generator.mappers import TableDataMapper
from core.invoice_generator.mappers.models import ResolvedTableData
from core.invoice_generator.mappers.transforms import extract_table_data
from core.invoice_generator.mappers.table import merge_static_content
from core.invoice_generator.mappers.footer import extract_summaries, format_pallet_counts


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


def test_helpers_merge_static_content():
    """Verify merge_static_content merges static values and resolves fallback descriptions."""
    data_rows = [{"col_item": "dynamic1"}, {"col_item": "dynamic2"}]
    static_content = {"col_static": ["Static1", "{col_desc_fallback}"]}
    dynamic_mapping_rules = {
        "col_desc": {
            "fallback": {
                "standard": "Standard Desc"
            }
        }
    }

    merge_static_content(
        data_rows=data_rows,
        static_content=static_content,
        dynamic_mapping_rules=dynamic_mapping_rules,
        DAF_mode=False,
        custom_mode=False
    )

    # Values should be merged into col_static
    assert data_rows[0]["col_static"] == "Static1"
    assert data_rows[1]["col_static"] == "Standard Desc"

    # Test extending rows if we have more static values than data rows
    data_rows_short = [{"col_item": "dynamic1"}]
    static_content_long = {"col_static": ["Static1", "Static2"]}
    merge_static_content(
        data_rows=data_rows_short,
        static_content=static_content_long,
        dynamic_mapping_rules={},
        DAF_mode=False,
        custom_mode=False
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
    l, w, p = extract_summaries(data_source_dict, footer_data={}, table_key=None)
    assert l == "l_sum"
    assert w == "w_sum"
    assert p == 4

    # Pre-calculated in footer_data grand_total
    footer_data = {
        "grand_total": {
            "col_pallet_count": 10
        }
    }
    _, _, p = extract_summaries(None, footer_data, table_key=None)
    assert p == 10

    # Multi-table table_totals list
    footer_data_multi = {
        "table_totals": [
            {"col_pallet_count": 3},
            {"col_pallet_count": 5}
        ]
    }
    _, _, p0 = extract_summaries(None, footer_data_multi, table_key=0)
    _, _, p1 = extract_summaries(None, footer_data_multi, table_key=1)
    assert p0 == 3
    assert p1 == 5


def test_helpers_format_pallet_counts():
    """Verify that format_pallet_counts formats counts and carries values forward correctly."""
    data_rows = [
        {"col_item": "Item 1", "col_pallet_count": 1},
        {"col_item": "Item 2", "col_pallet_count": 0},
        {"col_item": "Item 3", "col_pallet_count": 1}
    ]
    footer_data = {
        "grand_total": {
            "col_pallet_count": 5
        }
    }
    format_pallet_counts(
        data_rows=data_rows,
        num_data_rows=3,
        pallet_col_id="col_pallet_count",
        footer_data=footer_data,
        table_key=None
    )

    # 1-5, then carry 1-5 forward, then next is 2-5
    assert data_rows[0]["col_pallet_count"] == "1-5"
    assert data_rows[1]["col_pallet_count"] == "1-5"
    assert data_rows[2]["col_pallet_count"] == "2-5"
