import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.data import DataTableBuilderStyler
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_data_table_builder_parent_no_horizontal_merge():
    column_mapping = {
        "col_static": 1,
        "col_qty_header": 2,
        "col_qty_pcs": 2,
        "col_qty_sf": 3
    }
    column_colspan = {
        "col_static": 1,
        "col_qty_header": 1,  # Parent column gets 1
        "col_qty_pcs": 1,
        "col_qty_sf": 1
    }
    
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry, column_colspan=column_colspan)
    
    resolved_data = {
        "data_rows": [
            {
                1: "Item A",
                2: 10,   # PCS
                3: 100.5 # SF
            }
        ]
    }
    
    builder = DataTableBuilderStyler(
        grid=grid,
        resolved_data=resolved_data
    )
    builder.build()
    
    # Assert values in Grid
    assert grid.get_cell(-1, "col_static").value == "Item A"
    assert grid.get_cell(-1, "col_qty_pcs").value == 10
    assert grid.get_cell(-1, "col_qty_sf").value == 100.5
    
    # Assert that there is NO merge registered for the parent column on the data row
    assert grid.get_cell(-1, "col_qty_header").merge is None

def test_data_table_builder_vertical_merge_consecutive():
    column_mapping = {"col_po": 1, "col_qty": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    resolved_data = {
        "data_rows": [
            {1: "PO-001", 2: 10, "rowspans": {"col_po": 2}},
            {1: "PO-001", 2: 15, "rowspans": {"col_po": 0}},
            {1: "PO-002", 2: 20, "rowspans": {"col_po": 1}}
        ]
    }
    
    builder = DataTableBuilderStyler(
        grid=grid,
        resolved_data=resolved_data
    )
    builder.build()
    
    # PO-001 spans first two rows, should be merged
    cell1 = grid.get_cell(-3, "col_po")
    assert cell1.value == "PO-001"
    assert cell1.merge is not None
    assert cell1.merge.row_span == 2
    
    # PO-002 is distinct, no merge
    cell3 = grid.get_cell(-1, "col_po")
    assert cell3.value == "PO-002"
    assert cell3.merge is None


def test_data_table_builder_vertical_merge_skip_int():
    column_mapping = {"col_num": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    resolved_data = {
        "data_rows": [
            {1: "100", "rowspans": {"col_num": 1}},
            {1: "100", "rowspans": {"col_num": 1}}
        ]
    }
    
    builder = DataTableBuilderStyler(
        grid=grid,
        resolved_data=resolved_data
    )
    builder.build()
    
    # Numeric values should skip vertical merging
    cell = grid.get_cell(-2, "col_num")
    assert cell.merge is None


def test_apply_vertical_merges():
    from core.invoice_generator.utils.merge_transformer import apply_vertical_merges
    
    # Test merging of description (when uniform) and pallet numbers
    rows = [
        {"col_pallet_no": "P1", "col_desc": "Uniform Item"},
        {"col_pallet_no": "P1", "col_desc": "Uniform Item"},
        {"col_pallet_no": "P2", "col_desc": "Uniform Item"}
    ]
    
    result = apply_vertical_merges(rows)
    assert result[0]["rowspans"]["col_pallet_no"] == 2
    assert result[1]["rowspans"]["col_pallet_no"] == 0
    assert result[2]["rowspans"]["col_pallet_no"] == 1
    
    assert result[0]["rowspans"]["col_desc"] == 3
    assert result[1]["rowspans"]["col_desc"] == 0
    assert result[2]["rowspans"]["col_desc"] == 0

    # Test that digit-only strings in col_pallet_no are not merged
    rows_digits = [
        {"col_pallet_no": "1", "col_desc": "Uniform Item"},
        {"col_pallet_no": "1", "col_desc": "Uniform Item"}
    ]
    result_digits = apply_vertical_merges(rows_digits)
    assert result_digits[0]["rowspans"].get("col_pallet_no", 1) == 1
    assert result_digits[1]["rowspans"].get("col_pallet_no", 1) == 1


def test_data_table_builder_parent_logical_keys_no_overwrite():
    column_mapping = {
        "col_static": 1,
        "col_qty_header": 2,
        "col_qty_pcs": 2,
        "col_qty_sf": 3
    }
    column_colspan = {
        "col_static": 1,
        "col_qty_header": 1,
        "col_qty_pcs": 1,
        "col_qty_sf": 1
    }
    
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry, column_colspan=column_colspan)
    
    resolved_data = {
        "data_rows": [
            {
                "col_static": "Item A",
                "col_qty_pcs": 10,
                "col_qty_sf": 100.5
            }
        ]
    }
    
    builder = DataTableBuilderStyler(
        grid=grid,
        resolved_data=resolved_data,
        parent_column_ids=["col_qty_header"]
    )
    builder.build()
    
    # Assert values in Grid
    assert grid.get_cell(-1, "col_static").value == "Item A"
    assert grid.get_cell(-1, "col_qty_pcs").value == 10
    assert grid.get_cell(-1, "col_qty_sf").value == 100.5
