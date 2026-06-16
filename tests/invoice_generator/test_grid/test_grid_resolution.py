import pytest
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_resolve_column_by_logical_id():
    column_mapping = {"col_a": 1, "col_b": 4}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    assert grid.get_column_index("col_a") == 1
    assert grid.get_column_index("col_b") == 4
    assert grid.get_column_letter("col_a") == "A"
    assert grid.get_column_letter("col_b") == "D"

def test_resolve_column_by_integer_and_string_integer():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    # Integer inputs should resolve directly
    assert grid.get_column_index(5) == 5
    assert grid.get_column_letter(5) == "E"
    
    # String representation of integers should resolve
    assert grid.get_column_index("3") == 3
    assert grid.get_column_letter("3") == "C"

def test_resolve_column_not_found():
    column_mapping = {"col_a": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    with pytest.raises(ValueError, match="Column ID 'col_invalid' not found in grid mapping"):
        grid.get_column_index("col_invalid")
        
    with pytest.raises(ValueError, match="Column ID 'col_invalid' not found in grid mapping"):
        grid.get_column_letter("col_invalid")
