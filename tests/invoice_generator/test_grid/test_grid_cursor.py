import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_grid_initial_cursor_state():
    column_mapping = {"col_a": 1, "col_b": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    assert grid.start_row_index == 0
    assert grid._cursor_row == 0
    assert grid.num_columns == 2

def test_grid_set_start_row():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    grid.set_start_row(15)
    assert grid.start_row_index == 15

def test_grid_advance_row():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    grid.advance_row()
    assert grid._cursor_row == 1
    
    grid.advance_row(4)
    assert grid._cursor_row == 5

def test_grid_num_columns_empty():
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid({}, style_registry)
    assert grid.num_columns == 0
