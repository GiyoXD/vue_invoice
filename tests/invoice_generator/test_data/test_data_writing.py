import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.data import DataTableBuilderStyler
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_data_table_builder_error_propagation():
    column_mapping = {"col_no": 1, "col_desc": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    # Mock write to raise an exception
    def mock_write(*args, **kwargs):
        raise ValueError("Simulated write error")
    grid.write = mock_write
    
    builder = DataTableBuilderStyler(
        grid=grid,
        resolved_data={"data_rows": [{"col_no": 1}]}
    )
    
    with pytest.raises(ValueError, match="Simulated write error"):
        builder.build()

def test_data_table_builder_value_conversions():
    column_mapping = {"col_int": 1, "col_float": 2, "col_str": 3, "col_empty": 4}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    resolved_data = {
        "data_rows": [
            {
                1: "15",
                2: "25.5",
                3: "Hello",
                4: "   "
            }
        ]
    }
    
    builder = DataTableBuilderStyler(grid, resolved_data)
    builder.build()
    
    # Assert values converted and written correctly
    assert grid.get_cell(-1, "col_int").value == 15
    assert grid.get_cell(-1, "col_float").value == 25.5
    assert grid.get_cell(-1, "col_str").value == "Hello"
    assert grid.get_cell(-1, "col_empty").value is None

def test_data_table_builder_missing_col_no_generation():
    column_mapping = {"col_no": 1, "col_desc": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    # row_data only defines col_desc
    resolved_data = {
        "data_rows": [
            {2: "Item 1"},
            {2: "Item 2"}
        ]
    }
    
    builder = DataTableBuilderStyler(grid, resolved_data)
    builder.build()
    
    # Assert missing col_no is auto-generated based on loop index + 1
    assert grid.get_cell(-2, "col_no").value == 1
    assert grid.get_cell(-1, "col_no").value == 2
