import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.models.config.layout import FooterConfigModel

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
    grid.set_section_bounds("data", 0, 4) # mock data section
    
    rows = [[
        {"col_id": "col_po", "value": "TOTAL:", "style_context": "footer"},
        {"col_id": "col_item", "value": "10 PALLETS", "style_context": "footer"},
        {"col_id": "col_qty", "formula": "SUM", "target_section": "data", "style_context": "footer"}
    ]]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config,
        payload={"col_pallet_count": 10, "pallet_count": 10, "multiple": "S"}
    )
    builder.build()
    
    assert grid._grid[0][1].value == "TOTAL:"
    assert grid._grid[0][2].value == "10 PALLETS"
    assert grid._grid[0][3].value == "=SUM(C1:C5)" # col_qty resolves to index 3 (C)


def test_footer_builder_pallet_count_templating():
    column_mapping = {"col_item": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    rows = [[
        {"col_id": "col_item", "value": "{pallet_count} PALLET{multiple}", "style_context": "footer"}
    ]]
    footer_config = FooterConfigModel(rows=rows)
    

    
    builder1 = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config,
        payload={"col_pallet_count": 1, "multiple": ""}
    )
    builder1.build()
    assert grid._grid[0][1].value == "1 PALLET"
    
    # Test pluralization
    grid2 = Grid(column_mapping, style_registry)
    builder2 = TableFooterBuilder(
        grid=grid2,
        footer_config=footer_config,
        payload={"col_pallet_count": 5, "multiple": "S"}
    )
    builder2.build()
    assert grid2._grid[0][1].value == "5 PALLETS"

def test_footer_builder_main_footer_merges():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    rows = [[
        {"col_id": "col_po", "value": "TOTAL:", "colspan": 2, "style_context": "footer"}
    ]]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config
    )
    builder.build()
    
    cell = grid._grid[0][1]
    assert cell.value == "TOTAL:"
    assert cell.merge is not None
    assert cell.merge.min_col == 1
    assert cell.merge.max_col == 2
    assert cell.merge.row_span == 1

def test_footer_builder_pallet_count_zero_does_not_skip_row():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(1)
    
    rows = [[
        {"col_id": "col_po", "value": "TOTAL:", "style_context": "footer"},
        {"col_id": "col_item", "value": "{pallet_count} PALLET{multiple}", "style_context": "footer"},
        {"col_id": "col_qty", "value": "100", "style_context": "footer"}
    ]]
    footer_config = FooterConfigModel(rows=rows)
    builder = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config,
        payload={"col_pallet_count": 0, "pallet_count": 0, "multiple": ""}
    )
    builder.build()
    
    # Assert row is written (grid advances)
    assert grid._cursor_row == 1
    # TOTAL: and qty 100 should be written
    assert grid._grid[0][1].value == "TOTAL:"
    assert grid._grid[0][3].value == 100
    # Pallet count cell should be skipped (None), not written
    assert grid._grid[0][2].value is None

