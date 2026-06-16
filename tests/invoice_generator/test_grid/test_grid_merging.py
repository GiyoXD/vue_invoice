import pytest
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_grid_merge_cells_and_clear_span():
    column_mapping = {"col_a": 1, "col_b": 2, "col_c": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    # 1. Write values to cells first
    grid.write(0, "col_a", "TopLeft")
    grid.write(0, "col_b", "Overwritten1")
    grid.write(1, "col_a", "Overwritten2")
    grid.write(1, "col_b", "Overwritten3")
    
    # 2. Merge 2x2 area starting at (0, "col_a")
    grid.merge(row=0, col_id="col_a", rowspan=2, colspan=2)
    
    # Assert top-left cell preserves value and has TemplateMerge info
    top_left = grid.get_cell(0, "col_a")
    assert top_left.value == "TopLeft"
    assert top_left.merge is not None
    assert top_left.merge.min_col == 1
    assert top_left.merge.max_col == 2
    assert top_left.merge.row_span == 2
    
    # Assert other cells in the merge range are cleared
    assert grid.get_cell(0, "col_b").value is None
    assert grid.get_cell(1, "col_a").value is None
    assert grid.get_cell(1, "col_b").value is None

def test_grid_merge_noop():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.write(0, "col_a", "Value")
    
    # 1x1 merge should be a no-op
    grid.merge(row=0, col_id="col_a", rowspan=1, colspan=1)
    
    assert grid.get_cell(0, "col_a").merge is None
