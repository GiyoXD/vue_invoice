import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.models.config.layout import FooterConfigModel

def test_footer_orchestration_basic():
    column_mapping = {"col_po": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    rows = [[{"col_id": "col_po", "value": "TOTAL:", "style_context": "footer"}]]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config
    )
    builder.build()
    
    # Assert grid advanced by 1 row (regular footer)
    assert grid._cursor_row == 1
    # Query at relative offset -1 since cursor has advanced by 1
    assert grid.get_cell(-1, "col_po").value == "TOTAL:"


def test_footer_orchestration_empty_config_raises():
    column_mapping = {"col_po": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    footer_config = FooterConfigModel(rows=[])
    
    builder = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config
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
    
    footer_config = FooterConfigModel(rows=[])
    
    builder = TableFooterBuilder(
        grid=grid,
        footer_config=footer_config
    )
    
    # Assert builder resolves ranges dynamically from grid sections
    # Physical start: 10 + 2 = 12. Physical end: 10 + 5 = 15.
    assert builder.sum_ranges == [(12, 15)]
