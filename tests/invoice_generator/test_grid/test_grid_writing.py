import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_grid_cursor_offset():
    column_mapping = {"col_a": 1, "col_b": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10) # Start row in sheet

    # 1. Cursor is 0. Writes go to relative rows 0 and 1
    grid.write(0, "col_a", "val_a0")
    grid.write(1, "col_b", "val_b1")
    
    assert grid.get_cell(0, "col_a").value == "val_a0"
    assert grid.get_cell(1, "col_b").value == "val_b1"
    
    # 2. Advance cursor by 2
    grid.advance_row(2)
    assert grid._cursor_row == 2
    
    # Writes at relative index 0 now go to absolute grid index 2
    grid.write(0, "col_a", "val_a2")
    grid.write(1, "col_b", "val_b3")
    
    # Query relative to cursor
    assert grid.get_cell(0, "col_a").value == "val_a2"
    assert grid.get_cell(1, "col_b").value == "val_b3"
    
    # Raw matrix check to ensure physical row indices 0, 1, 2, 3 were written
    assert grid._grid[0][1].value == "val_a0"
    assert grid._grid[1][2].value == "val_b1"
    assert grid._grid[2][1].value == "val_a2"
    assert grid._grid[3][2].value == "val_b3"

def test_grid_write_formula_absolute_row_offset():
    column_mapping = {"col_a": 1, "col_b": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10) # Absolute worksheet starts at 10
    
    # Advance cursor by 5
    grid.advance_row(5)
    
    # Write formula at relative row 1
    # Absolute row in sheet should be: start_row(10) + cursor(5) + relative_row(1) = 16
    grid.write_formula(1, "col_b", "={col_ref_0}*10", ["col_a"])
    
    cell = grid.get_cell(1, "col_b")
    assert cell.value == "=A16*10"

def test_grid_write_with_context_style():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({
        "columns": {
            "col_a": {"format": "@"}
        },
        "row_contexts": {
            "header": {"bold": True}
        }
    })
    grid = Grid(column_mapping, style_registry)
    grid.write(0, "col_a", "Header Text", context="header")
    
    cell = grid.get_cell(0, "col_a")
    assert cell.value == "Header Text"
    # Verify some style attributes exist on cell style
    assert cell.style is not None
    assert cell.style.font.bold is True

def test_grid_get_row_models():
    column_mapping = {"col_a": 1, "col_b": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    grid.write(0, "col_a", "row0_a")
    grid.write(2, "col_b", "row2_b")
    
    rows = grid.get_row_models()
    assert len(rows) == 3 # rows at indices 0, 1, 2
    assert rows[0].relative_index == 0
    assert rows[0].cells[0].value == "row0_a"
    assert len(rows[1].cells) == 0 # Row 1 has no cells
    assert rows[2].relative_index == 2
    assert rows[2].cells[0].value == "row2_b"
