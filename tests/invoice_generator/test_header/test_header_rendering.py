import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.header import HeaderBuilderStyler
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_header_builder_parent_and_children():
    column_mapping = {
        "col_static": 1,
        "col_qty_header": 2,
        "col_qty_pcs": 2,
        "col_qty_sf": 3
    }
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10) # Absolute worksheet row starts at 10
    
    # Define parent and children headers structure
    bundled_columns = [
        {
            "id": "col_static",
            "header": "Mark & Nº",
            "rowspan": 2,
            "colspan": 1
        },
        {
            "id": "col_qty_header",
            "header": "Quantity(SF)",
            "children": [
                {
                    "id": "col_qty_pcs",
                    "header": "PCS"
                },
                {
                    "id": "col_qty_sf",
                    "header": "SF"
                }
            ]
        }
    ]
    
    builder = HeaderBuilderStyler(
        grid=grid,
        start_row=10,
        bundled_columns=bundled_columns
    )
    builder.build()
    
    # Assert section bounds are marked correctly
    assert grid._sections.get("header") == (0, 1) # 2 relative rows
    assert grid.get_section_range("header") == (10, 11) # Absolute rows 10 and 11
    
    # Assert values are written to Grid correctly
    assert grid._grid[0][1].value == "Mark & Nº"
    assert grid._grid[0][2].value == "Quantity(SF)"
    
    # Row 1 corresponds to grid._grid[1]
    assert grid._grid[1][1].value is None # Merged under col_static
    assert grid._grid[1][2].value == "PCS"
    assert grid._grid[1][3].value == "SF"
    
    # Assert merges are correctly registered
    merge_static = grid._grid[0][1].merge
    assert merge_static is not None
    assert merge_static.min_col == 1
    assert merge_static.max_col == 1
    assert merge_static.row_span == 2
    
    merge_qty = grid._grid[0][2].merge
    assert merge_qty is not None
    assert merge_qty.min_col == 2
    assert merge_qty.max_col == 3
    assert merge_qty.row_span == 1
    
    # Assert cursor advanced by number of header rows
    assert grid._cursor_row == 2

def test_header_builder_packing_list():
    bundled_columns = [
      {
        "id": "col_static",
        "header": "Mark & Nº",
        "rowspan": 2,
        "colspan": 1
      },
      {
        "id": "col_po",
        "header": "P.O Nº",
        "rowspan": 2
      },
      {
        "id": "col_item",
        "header": "ITEM Nº",
        "rowspan": 2
      },
      {
        "id": "col_desc",
        "header": "Description",
        "rowspan": 2
      },
      {
        "id": "col_qty_header",
        "header": "Quantity(SF)",
        "colspan": 2,
        "children": [
          {
            "id": "col_qty_pcs",
            "header": "PCS"
          },
          {
            "id": "col_qty_sf",
            "header": "SF"
          }
        ]
      },
      {
        "id": "col_net",
        "header": "N.W (kgs)",
        "rowspan": 2
      },
      {
        "id": "col_gross",
        "header": "G.W (kgs)",
        "rowspan": 2
      },
      {
        "id": "col_cbm",
        "header": "CBM",
        "rowspan": 2
      }
    ]
    
    column_mapping = {
        "col_static": 1,
        "col_po": 2,
        "col_item": 3,
        "col_desc": 4,
        "col_qty_header": 5,
        "col_qty_pcs": 5,
        "col_qty_sf": 6,
        "col_net": 7,
        "col_gross": 8,
        "col_cbm": 9
    }
    
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_start_row(10)
    
    builder = HeaderBuilderStyler(
        grid=grid,
        start_row=10,
        bundled_columns=bundled_columns
    )
    builder.build()
    
    # Assert values in Grid
    assert grid._grid[0][1].value == "Mark & Nº"
    assert grid._grid[0][2].value == "P.O Nº"
    assert grid._grid[0][3].value == "ITEM Nº"
    assert grid._grid[0][4].value == "Description"
    
    assert grid._grid[0][5].value == "Quantity(SF)"
    assert grid._grid[1][5].value == "PCS"
    assert grid._grid[1][6].value == "SF"
    
    assert grid._grid[0][7].value == "N.W (kgs)"
    
    # Assert merges are correctly registered
    merge_qty = grid._grid[0][5].merge
    assert merge_qty is not None
    assert merge_qty.min_col == 5
    assert merge_qty.max_col == 6
    assert merge_qty.row_span == 1
