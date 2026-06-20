import pytest
from typing import Dict, Any

from core.invoice_generator.config.table_value_adapter import TableDataAdapter
from core.invoice_generator.models.table_adapter import ResolvedTableData, StaticInfoModel
from core.invoice_generator.config.table_value_adapter.helpers import (
    extract_table_data,
    merge_static_content,
    extract_summaries,
    format_pallet_counts
)


def test_resolved_table_data_model_dict_compat():
    """Verify that ResolvedTableData supports dictionary-like access for backward compatibility."""
    model = ResolvedTableData(
        data_rows=[{1: "row1"}],
        pallet_counts=[2],
        num_data_rows=1,
        static_info=StaticInfoModel(col1_index=0, num_static_labels=1),
        formula_rules={"col_total": "SUM"},
        static_content={"col_static": ["Static Value"]},
        leather_summary={"BUFFALO": {"col_pallet_count": 2}},
        weight_summary={"net": 100.0, "gross": 110.0},
        pallet_summary_total=5
    )

    # __getitem__ compatibility
    assert model["data_rows"] == [{1: "row1"}]
    assert model["pallet_counts"] == [2]
    assert model["num_data_rows"] == 1
    assert model["pallet_summary_total"] == 5
    with pytest.raises(KeyError):
        _ = model["non_existent_key"]

    # get compatibility
    assert model.get("data_rows") == [{1: "row1"}]
    assert model.get("pallet_summary_total", 0) == 5
    assert model.get("non_existent_key", "default_val") == "default_val"
    assert model.get("non_existent_key") is None

    # __contains__ compatibility
    assert "data_rows" in model
    assert "pallet_summary_total" in model
    assert "non_existent_key" not in model


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
    data_rows = [{1: "dynamic1"}, {1: "dynamic2"}]
    static_content = {"col_static": ["Static1", "{col_desc_fallback}"]}
    column_id_map = {"col_static": 2}
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
        column_id_map=column_id_map,
        dynamic_mapping_rules=dynamic_mapping_rules,
        DAF_mode=False,
        custom_mode=False
    )

    # Values should be merged into col_static (index 2)
    assert data_rows[0][2] == "Static1"
    assert data_rows[1][2] == "Standard Desc"

    # Test extending rows if we have more static values than data rows
    data_rows_short = [{1: "dynamic1"}]
    static_content_long = {"col_static": ["Static1", "Static2"]}
    merge_static_content(
        data_rows=data_rows_short,
        static_content=static_content_long,
        column_id_map=column_id_map,
        dynamic_mapping_rules={},
        DAF_mode=False,
        custom_mode=False
    )
    assert len(data_rows_short) == 2
    assert data_rows_short[0][2] == "Static1"
    assert data_rows_short[1][2] == "Static2"


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
        {1: "Item 1", 2: 1},  # pallet count col index 2
        {1: "Item 2", 2: 0},
        {1: "Item 3", 2: 1}
    ]
    footer_data = {
        "grand_total": {
            "col_pallet_count": 5
        }
    }
    format_pallet_counts(
        data_rows=data_rows,
        num_data_rows=3,
        pallet_col_idx=2,
        footer_data=footer_data,
        table_key=None
    )

    # 1-5, then carry 1-5 forward, then next is 2-5
    assert data_rows[0][2] == "1-5"
    assert data_rows[1][2] == "1-5"
    assert data_rows[2][2] == "2-5"
