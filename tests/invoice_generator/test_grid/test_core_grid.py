import pytest
from core.models.grid import Grid
from core.models.cell import CellStyle, FontStyle, AlignmentStyle

def test_core_grid_basic_write_read():
    grid = Grid()
    
    # Write some physical values
    grid.write(0, 1, "A1")
    grid.write(1, 2, "B2")
    grid.set_row_height(0, 15.0)
    grid.set_row_height(1, 20.0)
    
    # Read back values
    assert grid.get_cell(0, 1).value == "A1"
    assert grid.get_cell(1, 2).value == "B2"
    assert grid.get_cell(0, 2).value is None
    
    assert grid.get_row_height(0) == 15.0
    assert grid.get_row_height(1) == 20.0

def test_core_grid_merge_resolution():
    grid = Grid()
    
    grid.write(2, 2, "Merged Parent")
    # Merge block: row 2 to 3, column 2 to 4
    grid.merge(row=2, col=2, rowspan=2, colspan=3)
    
    # The parent cell holds the merge info and the value
    parent_cell = grid.get_cell(2, 2)
    assert parent_cell.value == "Merged Parent"
    assert parent_cell.merge is not None
    assert parent_cell.merge.min_col == 2
    assert parent_cell.merge.max_col == 4
    assert parent_cell.merge.row_span == 2
    
    # Queries to other cells in the merge block must automatically resolve to parent
    assert grid.get_cell(2, 3, resolve_merge=True).value == "Merged Parent"
    assert grid.get_cell(3, 4, resolve_merge=True).value == "Merged Parent"
    
    # Cells outside the merge block do not resolve to parent
    assert grid.get_cell(2, 5).value is None
    assert grid.get_cell(4, 2).value is None

def test_core_grid_to_dict_and_from_dict():
    grid = Grid()
    
    # Add values, styling, row heights, and merges
    font = FontStyle(name="Arial", size=11, bold=True)
    align = AlignmentStyle(horizontal="center", vertical="center", wrap_text=True)
    style = CellStyle(font=font, alignment=align)
    
    grid.write(0, 1, "Header", style=style)
    grid.write(1, 1, "Data1")
    grid.write(1, 2, "Data2")
    grid.merge(1, 1, rowspan=1, colspan=2)
    grid.set_row_height(0, 18.0)
    
    # Serialize
    d = grid.to_dict()
    
    # Reconstruct
    grid_rebuilt = Grid.from_dict(d)
    
    # Verify rebuilt grid values
    assert grid_rebuilt.get_row_height(0) == 18.0
    assert grid_rebuilt.get_cell(0, 1).value == "Header"
    assert grid_rebuilt.get_cell(0, 1).style.font.name == "Arial"
    assert grid_rebuilt.get_cell(0, 1).style.font.bold is True
    assert grid_rebuilt.get_cell(0, 1).style.alignment.horizontal == "center"
    
    # Verify merges rebuilt
    assert grid_rebuilt.get_cell(1, 2, resolve_merge=True).value == "Data1"  # Resolves through merge map
    parent_merge = grid_rebuilt.get_cell(1, 1).merge
    assert parent_merge.min_col == 1
    assert parent_merge.max_col == 2
    assert parent_merge.row_span == 1

def test_core_grid_get_row_models():
    grid = Grid()
    grid.write(0, 1, "Row0 Col1")
    grid.write(2, 5, "Row2 Col5")
    
    row_models = grid.get_row_models()
    assert len(row_models) == 3
    assert row_models[0].relative_index == 0
    assert row_models[0].cells[0].value == "Row0 Col1"
    assert row_models[2].relative_index == 2
    assert row_models[2].cells[0].value == "Row2 Col5"
