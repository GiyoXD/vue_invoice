import pytest
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_section_bounds_tracking():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10)
    
    # 1. Advance row to 2, then mark start
    grid.advance_row(2)
    grid.mark_section_start("data")
    
    # 2. Advance row by 3 (simulating writing 3 rows), then mark end
    grid.advance_row(3)
    grid.mark_section_end("data")
    
    # Grid cursor is now at 5. Relative range is (2, 4)
    # Physical range should be: start_row(10) + start(2) = 12 to start_row(10) + end(4) = 14
    start, end = grid.get_section_range("data")
    assert start == 12
    assert end == 14

def test_section_bounds_fallback():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10)
    grid.advance_row(3)
    
    # Query section range for non-existent section should fallback to current cursor
    # start_row(10) + cursor(3) = 13
    start, end = grid.get_section_range("non_existent")
    assert start == 13
    assert end == 13
