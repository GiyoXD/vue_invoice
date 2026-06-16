import pytest
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.styling.models import FooterData

def test_footer_builder_pads_styles_without_erasing_values():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({
        "columns": {
            "col_po": {"format": "@"},
            "col_item": {"format": "@"},
            "col_qty": {"format": "#,##0"}
        },
        "row_contexts": {
            "footer": {"bold": True}
        }
    })
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    footer_config = {
        "footer_cells": [
            ["TOTAL:", "col_po"],
            ["10 PALLETS", "col_item"]
        ],
        "sum_cols": ["col_qty"],
        "merge_rules": []
    }
    
    footer_data = FooterData(
        footer_row_start_idx=1,
        data_start_row=1,
        data_end_row=5,
        total_pallets=10,
        weight_summary=None,
        leather_summary=None
    )
    
    data_config = {
        "sum_ranges": [(1, 5)],
        "footer_config": footer_config
    }
    
    builder = TableFooterBuilder(
        grid=grid,
        footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={"pallet_count": 10},
        data_config=data_config
    )
    
    builder._build_main_footer(row=0, footer_type="regular")
    
    assert grid.get_cell(0, "col_po").value == "TOTAL:"
    assert grid.get_cell(0, "col_item").value == "10 PALLETS"
    assert grid.get_cell(0, "col_qty").value == "=SUM(C1:C5)" # col_qty resolves to index 3 (C)


def test_footer_builder_pallet_count_templating():
    column_mapping = {"col_item": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    footer_config = {
        "footer_cells": [
            ["{pallet_count} PALLET{multiple}", "col_item"]
        ],
        "sum_cols": [],
        "merge_rules": []
    }
    
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=1,
        weight_summary=None, leather_summary=None
    )
    
    builder1 = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={"pallet_count": 1},
        data_config={"sum_ranges": [(1, 5)], "footer_config": footer_config}
    )
    builder1._build_main_footer(row=0, footer_type="regular")
    assert grid.get_cell(0, "col_item").value == "1 PALLET"
    
    # Test pluralization
    grid2 = Grid(column_mapping, style_registry)
    builder2 = TableFooterBuilder(
        grid=grid2, footer_data=footer_data,
        style_config={"styling_config": style_registry},
        context_config={"pallet_count": 5},
        data_config={"sum_ranges": [(1, 5)], "footer_config": footer_config}
    )
    builder2._build_main_footer(row=0, footer_type="regular")
    assert grid2.get_cell(0, "col_item").value == "5 PALLETS"

def test_footer_builder_main_footer_merges():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    footer_config = {
        "footer_cells": [["TOTAL:", "col_po"]],
        "sum_cols": [],
        "merge_rules": [
            {"start_column_id": "col_po", "colspan": 2}
        ]
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
    
    builder._build_main_footer(row=0, footer_type="regular")
    
    cell = grid.get_cell(0, "col_po")
    assert cell.value == "TOTAL:"
    assert cell.merge is not None
    assert cell.merge.min_col == 1
    assert cell.merge.max_col == 2
    assert cell.merge.row_span == 1
