import pytest
from openpyxl import Workbook
from core.invoice_generator.builders.table.builder import TableBuilder

def test_resolve_columns_packing_list():
    wb = Workbook()
    ws = wb.active
    
    # Setup columns configuration exactly as in Packing List of JF_bundle_config.json
    sheet_config = {
        "structure": {
            "columns": [
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
        }
    }
    
    builder = TableBuilder(
        workbook=wb,
        worksheet=ws,
        style_config={},
        context_config={"sheet_name": "Packing list"},
        layout_config={"sheet_config": sheet_config}
    )
    
    bundled_columns, column_mapping, column_colspan = builder._resolve_columns()
    
    # Assert physical column index mappings
    assert column_mapping["col_static"] == 1
    assert column_mapping["col_po"] == 2
    assert column_mapping["col_item"] == 3
    assert column_mapping["col_desc"] == 4
    assert column_mapping["col_qty_header"] == 5
    assert column_mapping["col_qty_pcs"] == 5
    assert column_mapping["col_qty_sf"] == 6
    assert column_mapping["col_net"] == 7
    assert column_mapping["col_gross"] == 8
    assert column_mapping["col_cbm"] == 9
    
    # Assert parent column has colspan 1 in resolved column_colspan (preventing data-row horizontal merge)
    assert column_colspan["col_qty_header"] == 1
    assert column_colspan["col_qty_pcs"] == 1
    assert column_colspan["col_qty_sf"] == 1

def test_resolve_columns_daf_mode_filtering():
    wb = Workbook()
    ws = wb.active
    
    class ArgsMock:
        DAF = True
        custom = False
        
    sheet_config = {
        "structure": {
            "columns": [
              {"id": "col_a", "colspan": 1},
              {"id": "col_b", "colspan": 1, "skip_in_daf": True},
              {"id": "col_c", "colspan": 1}
            ]
        }
    }
    
    builder = TableBuilder(
        workbook=wb,
        worksheet=ws,
        style_config={},
        context_config={"sheet_name": "Test Sheet", "args": ArgsMock()},
        layout_config={"sheet_config": sheet_config}
    )
    
    bundled_columns, column_mapping, column_colspan = builder._resolve_columns()
    
    # Assert columns to skip are excluded from bundled_columns
    assert len(bundled_columns) == 2
    assert bundled_columns[0]["id"] == "col_a"
    assert bundled_columns[1]["id"] == "col_c"
    
    # Assert physical mapping offsets (col_c maps to physical column index 2)
    assert column_mapping["col_a"] == 1
    assert "col_b" not in column_mapping
    assert column_mapping["col_c"] == 2
