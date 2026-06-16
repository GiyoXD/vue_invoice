import pytest
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.models.footer import FooterData

def test_before_footer_addon():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=10,
        weight_summary=None, leather_summary=None
    )
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={},
        data_config={}
    )
    
    addon_config = {
        "column_id": "col_po",
        "text": "Special Remarks",
        "merge": 2
    }
    
    builder._build_before_footer(row=0, config=addon_config)
    
    cell = grid.get_cell(0, "col_po")
    assert cell.value == "Special Remarks"
    assert cell.merge is not None
    assert cell.merge.min_col == 1
    assert cell.merge.max_col == 2
    assert cell.merge.row_span == 1

def test_weight_summary_addon():
    column_mapping = {"col_desc": 1, "col_qty": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    weight_data = {"net": 500.5, "gross": 520.8}
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=10,
        weight_summary=weight_data, leather_summary=None
    )
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={},
        data_config={}
    )
    
    addon_config = {
        "label_col_id": "col_desc",
        "value_col_id": "col_qty"
    }
    
    builder._build_weight_summary_addon(row=0, config=addon_config)
    
    # Assert NW Row
    assert grid.get_cell(0, "col_desc").value == "NW(KGS)"
    assert grid.get_cell(0, "col_qty").value == 500.5
    
    # Assert GW Row
    assert grid.get_cell(1, "col_desc").value == "GW(KGS):"
    assert grid.get_cell(1, "col_qty").value == 520.8

def test_leather_summary_addon():
    column_mapping = {"col_desc": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    leather_data = {
        "BUFFALO": {"pallet_count": 2, "col_qty": 100},
        "COW": {"pallet_count": 3, "col_qty": 150}
    }
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=5,
        weight_summary=None, leather_summary=leather_data
    )
    
    # Leather summary addon only generates if sheet name is "Packing list"
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={"sheet_name": "Packing list"},
        data_config={"footer_config": {"sum_cols": ["col_qty"]}}
    )
    
    builder._build_leather_summary_addon(row=0, config={"enabled": True})
    
    # Assert Buffalo row (row 0)
    assert grid.get_cell(0, "col_desc").value == "BUFFALO LEATHER"
    assert grid.get_cell(0, "col_pallet_count").value == "2"
    assert grid.get_cell(0, "col_qty").value == 100
    
    # Assert Cow row (row 1)
    assert grid.get_cell(1, "col_desc").value == "LEATHER"
    assert grid.get_cell(1, "col_pallet_count").value == "3"
    assert grid.get_cell(1, "col_qty").value == 150
