import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.models.config.layout import FooterConfigModel

def test_before_footer_addon():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    rows = [[
        {
            "col_id": "col_po",
            "value": "Special Remarks",
            "colspan": 2,
            "style_context": "footer"
        }
    ]]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        payload={"pallet_count": 10, "multiple": "S"},
        footer_config=footer_config
    )
    builder.build()
    
    cell = grid._grid[0][1]
    assert cell.value == "Special Remarks"
    assert cell.merge is not None
    assert cell.merge.min_col == 1
    assert cell.merge.max_col == 2
    assert cell.merge.row_span == 1

def test_weight_summary_addon():
    column_mapping = {"col_desc": 1, "col_qty": 2}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    rows = [
        [
            {"col_id": "col_desc", "value": "NW(KGS)", "style_context": "summary"},
            {"col_id": "col_qty", "value": "{weight_net}", "style_context": "summary"}
        ],
        [
            {"col_id": "col_desc", "value": "GW(KGS):", "style_context": "summary"},
            {"col_id": "col_qty", "value": "{weight_gross}", "style_context": "summary"}
        ]
    ]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        payload={"weight_net": 500.5, "weight_gross": 520.8},
        footer_config=footer_config
    )
    builder.build()
    
    assert grid._grid[0][1].value == "NW(KGS)"
    assert grid._grid[0][2].value == 500.5
    
    assert grid._grid[1][1].value == "GW(KGS):"
    assert grid._grid[1][2].value == 520.8

def test_leather_summary_addon():
    column_mapping = {"col_desc": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    rows = [
        [
            {"col_id": "col_desc", "value": "BUFFALO LEATHER", "style_context": "summary", "addon_type": "leather", "leather_key": "BUFFALO"},
            {"col_id": "col_pallet_count", "value": "{buffalo_pallet_count}", "style_context": "summary", "addon_type": "leather", "leather_key": "BUFFALO", "is_pallet": True},
            {"col_id": "col_qty", "value": "{buffalo_col_qty}", "style_context": "summary", "addon_type": "leather", "leather_key": "BUFFALO"}
        ],
        [
            {"col_id": "col_desc", "value": "LEATHER", "style_context": "summary", "addon_type": "leather", "leather_key": "COW"},
            {"col_id": "col_pallet_count", "value": "{cow_pallet_count}", "style_context": "summary", "addon_type": "leather", "leather_key": "COW", "is_pallet": True},
            {"col_id": "col_qty", "value": "{cow_col_qty}", "style_context": "summary", "addon_type": "leather", "leather_key": "COW"}
        ]
    ]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        payload={"buffalo_pallet_count": 2, "buffalo_col_qty": 100, "cow_pallet_count": 3, "cow_col_qty": 150},
        footer_config=footer_config
    )
    builder.build()
    
    assert grid._grid[0][1].value == "BUFFALO LEATHER"
    assert grid._grid[0][2].value == 2
    assert grid._grid[0][3].value == 100
    
    assert grid._grid[1][1].value == "LEATHER"
    assert grid._grid[1][2].value == 3
    assert grid._grid[1][3].value == 150

def test_declarative_rows_schema():
    column_mapping = {"col_desc": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_section_bounds("data", 5, 10)
    
    declarative_rows = [
        [
            {"col_id": "col_desc", "value": "VAT (7%):", "style_context": "footer"},
            {"col_id": "col_qty", "formula": "SUM", "target_section": "data", "style_context": "footer"},
            {"col_id": "col_pallet_count", "value": "{pallet_count} PALLETS", "style_context": "footer"}
        ]
    ]
    
    footer_config = FooterConfigModel(rows=declarative_rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        payload={"col_pallet_count": 8, "pallet_count": 8, "multiple": "S"},
        footer_config=footer_config
    )
    
    builder.build()
    
    assert grid._grid[0][1].value == "VAT (7%):"
    assert grid._grid[0][2].value == "8 PALLETS"
    assert grid._grid[0][3].value == "=SUM(C5:C10)"


def test_leather_summary_addon_auto_lookup():
    column_mapping = {"col_desc": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    rows = [
        [
            {"col_id": "col_desc", "value": "BUFFALO LEATHER", "style_context": "summary", "addon_type": "leather", "leather_key": "BUFFALO"},
            {"col_id": "col_pallet_count", "style_context": "summary", "addon_type": "leather", "leather_key": "BUFFALO"},
            {"col_id": "col_qty", "style_context": "summary", "addon_type": "leather", "leather_key": "BUFFALO"}
        ],
        [
            {"col_id": "col_desc", "value": "LEATHER", "style_context": "summary", "addon_type": "leather", "leather_key": "COW"},
            {"col_id": "col_pallet_count", "style_context": "summary", "addon_type": "leather", "leather_key": "COW"},
            {"col_id": "col_qty", "style_context": "summary", "addon_type": "leather", "leather_key": "COW"}
        ]
    ]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid,
        payload={"buffalo_col_pallet_count": 2, "buffalo_col_qty": 100, "cow_col_pallet_count": 3, "cow_col_qty": 150},
        footer_config=footer_config
    )
    builder.build()
    
    assert grid._grid[0][1].value == "BUFFALO LEATHER"
    assert grid._grid[0][2].value == 2
    assert grid._grid[0][3].value == 100
    
    assert grid._grid[1][1].value == "LEATHER"
    assert grid._grid[1][2].value == 3
    assert grid._grid[1][3].value == 150

