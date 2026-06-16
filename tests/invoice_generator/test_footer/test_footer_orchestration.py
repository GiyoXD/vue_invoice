import pytest
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.styling.models import FooterData

def test_footer_orchestration_basic():
    column_mapping = {"col_po": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    footer_config = {
        "footer_cells": [["TOTAL:", "col_po"]],
        "sum_cols": [],
        "merge_rules": []
    }
    
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=10,
        weight_summary=None, leather_summary=None
    )
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={},
        data_config={"footer_config": footer_config}
    )
    
    builder.build()
    
    # Assert grid advanced by 1 row (regular footer)
    assert grid._cursor_row == 1
    # Query at relative offset -1 since cursor has advanced by 1
    assert grid.get_cell(-1, "col_po").value == "TOTAL:"

def test_footer_orchestration_add_blank_before():
    column_mapping = {"col_po": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    footer_config = {
        "add_blank_before": True,
        "footer_cells": [["TOTAL:", "col_po"]],
        "sum_cols": [],
        "merge_rules": []
    }
    
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=10,
        weight_summary=None, leather_summary=None
    )
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={},
        data_config={"footer_config": footer_config}
    )
    
    builder.build()
    
    # Assert grid advanced by 2 rows (1 blank row, 1 footer row)
    assert grid._cursor_row == 2
    # Row 0 (relative -2) is the blank padded row, Row 1 (relative -1) is the main footer row
    assert grid.get_cell(-2, "col_po").value is None
    assert grid.get_cell(-1, "col_po").value == "TOTAL:"



def test_footer_orchestration_empty_config_raises():
    column_mapping = {"col_po": 1}
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
        data_config={"footer_config": {}}
    )
    
    with pytest.raises(ValueError, match="CANNOT BUILD FOOTER"):
        builder.build()

def test_footer_orchestration_dynamic_sum_ranges():
    column_mapping = {"col_qty": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10)
    
    # Simulate data section bounds
    grid.advance_row(2) # Cursor at 2
    grid.mark_section_start("data")
    grid.advance_row(4) # Cursor at 6
    grid.mark_section_end("data")
    
    footer_config = {
        "footer_cells": [],
        "sum_cols": ["col_qty"],
        "merge_rules": []
    }
    
    footer_data = FooterData(
        footer_row_start_idx=6, data_start_row=12, data_end_row=15, total_pallets=10,
        weight_summary=None, leather_summary=None
    )
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={},
        data_config={"footer_config": footer_config}
    )
    
    # Assert builder resolves ranges dynamically from grid sections
    # Physical start: 10 + 2 = 12. Physical end: 10 + 5 = 15.
    assert builder.sum_ranges == [(12, 15)]
