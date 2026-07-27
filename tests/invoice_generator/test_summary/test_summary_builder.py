import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.summary import SummaryBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.models.config.layout import FooterConfigModel

def test_summary_builder_pallet_count_templating():
    column_mapping = {"col_item": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    rows = [[
        {"col_id": "col_item", "value": "{col_pallet_count} PALLET{multiple}", "style_context": "summary"}
    ]]
    summary_config = FooterConfigModel(rows=rows)
    
    payload = {"col_pallet_count": 1, "multiple": ""}
    builder1 = SummaryBuilder(
        grid=grid,
        summary_config=summary_config,
        payload=payload
    )
    builder1.build()
    assert grid._grid[0][1].value == "1 PALLET"
    
    # Test pluralization
    grid2 = Grid(column_mapping, style_registry)
    builder2 = SummaryBuilder(
        grid=grid2,
        summary_config=summary_config,
        payload={"col_pallet_count": 5, "multiple": "S"}
    )
    builder2.build()
    assert grid2._grid[0][1].value == "5 PALLETS"



def test_summary_builder_rendering_and_merges():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    grid.set_section_bounds("data", 0, 4)  # mock data section
    
    rows = [
        [
            {"col_id": "col_po", "value": "TOTAL SUMMARY:", "colspan": 2, "style_context": "summary"},
            {"col_id": "col_qty", "formula": "SUM", "target_section": "data", "style_context": "summary"}
        ]
    ]
    summary_config = FooterConfigModel(rows=rows)
    
    builder = SummaryBuilder(
        grid=grid,
        summary_config=summary_config
    )
    builder.build()
    
    # Check value resolution
    cell_po = grid._grid[0][1]
    assert cell_po.value == "TOTAL SUMMARY:"
    assert cell_po.merge is not None
    assert cell_po.merge.min_col == 1
    assert cell_po.merge.max_col == 2
    
    # Check formula resolution
    cell_qty = grid._grid[0][3]
    assert cell_qty.value == "=SUM(C1:C5)"


def test_summary_builder_repeating_rows():
    column_mapping = {"col_leather_type": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    # repeating row configuration
    rows = [
        {
            "source_list": "leather_summary",
            "cells": [
                {"col_id": "col_leather_type", "style_context": "summary"},
                {"col_id": "col_pallet_count", "style_context": "summary"},
                {"col_id": "col_qty", "style_context": "summary"}
            ]
        }
    ]
    summary_config = FooterConfigModel(rows=rows)
    
    payload = {
        "leather_summary": [
            {"col_leather_type": "NAPPA", "col_pallet_count": 2, "col_qty": 100},
            {"col_leather_type": "SUEDE", "col_pallet_count": 1, "col_qty": 50}
        ]
    }
    
    builder = SummaryBuilder(
        grid=grid,
        summary_config=summary_config,
        payload=payload
    )
    builder.build()
    
    # Verify both records rendered
    assert grid._cursor_row == 2
    
    assert grid._grid[0][1].value == "NAPPA"
    assert grid._grid[0][2].value == 2
    assert grid._grid[0][3].value == 100
    
    assert grid._grid[1][1].value == "SUEDE"
    assert grid._grid[1][2].value == 1
    assert grid._grid[1][3].value == 50


def test_summary_builder_bounds_registration():
    column_mapping = {"col_po": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    rows = [
        [{"col_id": "col_po", "value": "SUMMARY", "style_context": "summary"}],
        {
            "source_list": "leather_summary",
            "cells": [{"col_id": "col_po", "style_context": "summary"}]
        }
    ]
    summary_config = FooterConfigModel(rows=rows)
    payload = {
        "leather_summary": [{"col_po": "ADDON"}]
    }
    
    builder = SummaryBuilder(
        grid=grid,
        summary_config=summary_config,
        payload=payload
    )
    builder.build()
    
    # Assert section bounds were correctly registered on the grid
    assert "summary" in grid._sections
    assert grid._sections["summary"] == (0, 1)
