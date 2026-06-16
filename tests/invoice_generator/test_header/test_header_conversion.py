import pytest
from core.invoice_generator.builders.table.header import HeaderBuilderStyler
from core.invoice_generator.builders.table.grid import Grid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_convert_bundled_columns_flat():
    column_mapping = {"col_po": 1, "col_item": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    bundled_columns = [
        {"id": "col_po", "header": "P.O №", "rowspan": 2},
        {"id": "col_item", "header": "ITEM", "colspan": 2}
    ]
    
    builder = HeaderBuilderStyler(grid, start_row=10, bundled_columns=bundled_columns)
    
    # Assert layout configs
    layout = builder.header_layout_config
    assert len(layout) == 2
    
    assert layout[0] == {
        'row': 0,
        'col': 0,
        'text': "P.O №",
        'id': "col_po",
        'rowspan': 2,
        'colspan': 1
    }
    
    assert layout[1] == {
        'row': 0,
        'col': 1,
        'text': "ITEM",
        'id': "col_item",
        'rowspan': 1,
        'colspan': 2
    }

def test_convert_bundled_columns_nested():
    column_mapping = {"col_static": 1, "col_qty_header": 2, "col_qty_pcs": 2, "col_qty_sf": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    bundled_columns = [
        {"id": "col_static", "header": "Mark & Nº", "rowspan": 2},
        {
            "id": "col_qty_header",
            "header": "Quantity(SF)",
            "children": [
                {"id": "col_qty_pcs", "header": "PCS"},
                {"id": "col_qty_sf", "header": "SF"}
            ]
        }
    ]
    
    builder = HeaderBuilderStyler(grid, start_row=10, bundled_columns=bundled_columns)
    layout = builder.header_layout_config
    assert len(layout) == 4 # 1 static parent, 1 nested parent, 2 nested children
    
    # Parent static
    assert layout[0] == {
        'row': 0,
        'col': 0,
        'text': "Mark & Nº",
        'id': "col_static",
        'rowspan': 2,
        'colspan': 1
    }
    
    # Parent quantity header
    assert layout[1] == {
        'row': 0,
        'col': 1,
        'text': "Quantity(SF)",
        'id': "col_qty_header",
        'rowspan': 1,
        'colspan': 2
    }
    
    # Child PCS
    assert layout[2] == {
        'row': 1,
        'col': 1,
        'text': "PCS",
        'id': "col_qty_pcs",
        'rowspan': 1,
        'colspan': 1
    }

    # Child SF
    assert layout[3] == {
        'row': 1,
        'col': 2,
        'text': "SF",
        'id': "col_qty_sf",
        'rowspan': 1,
        'colspan': 1
    }


def test_header_builder_no_bundled_columns_raises():
    column_mapping = {"col_po": 1}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    with pytest.raises(ValueError, match="No bundled columns provided"):
        HeaderBuilderStyler(grid, start_row=10, bundled_columns=[])
