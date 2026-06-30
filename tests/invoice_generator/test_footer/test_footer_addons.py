import pytest
from core.invoice_generator.builders.table.table_grid import Grid
from core.invoice_generator.builders.table.footer import TableFooterBuilder
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.models.footer import FooterData
from core.invoice_generator.models.config.layout import FooterConfigModel

def test_before_footer_addon():
    column_mapping = {"col_po": 1, "col_item": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=10,
        weight_summary=None, leather_summary=None
    )
    
    # Declarative config for before footer
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
        grid=grid, footer_data=footer_data,
        footer_config=footer_config
    )
    builder.build()
    
    # Grid cursor gets advanced. The written rows are at index 0.
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
    
    weight_data = {"net": 500.5, "gross": 520.8}
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=10,
        weight_summary=weight_data, leather_summary=None
    )
    
    # Declarative config for weight summary NW / GW rows
    rows = [
        [
            {"col_id": "col_desc", "value": "NW(KGS)", "style_context": "footer_addon"},
            {"col_id": "col_qty", "value": "{weight_net}", "style_context": "footer_addon"}
        ],
        [
            {"col_id": "col_desc", "value": "GW(KGS):", "style_context": "footer_addon"},
            {"col_id": "col_qty", "value": "{weight_gross}", "style_context": "footer_addon"}
        ]
    ]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        footer_config=footer_config
    )
    builder.build()
    
    # Assert NW Row (row 0)
    assert grid._grid[0][1].value == "NW(KGS)"
    assert grid._grid[0][2].value == 500.5
    
    # Assert GW Row (row 1)
    assert grid._grid[1][1].value == "GW(KGS):"
    assert grid._grid[1][2].value == 520.8

def test_leather_summary_addon():
    column_mapping = {"col_desc": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    
    leather_data = {
        "BUFFALO": {"pallet_count": 2, "col_qty": 100},
        "COW": {"pallet_count": 3, "col_qty": 150}
    }
    footer_data = FooterData(
        footer_row_start_idx=1, data_start_row=1, data_end_row=5, total_pallets=5,
        weight_summary=None, leather_summary=leather_data
    )
    
    # Declarative config for leather summary
    rows = [
        [
            {"col_id": "col_desc", "value": "BUFFALO LEATHER", "style_context": "footer_addon", "addon_type": "leather", "leather_key": "BUFFALO"},
            {"col_id": "col_pallet_count", "value": "{buffalo_pallet_count}", "style_context": "footer_addon", "addon_type": "leather", "leather_key": "BUFFALO", "is_pallet": True},
            {"col_id": "col_qty", "value": "{buffalo_col_qty}", "style_context": "footer_addon", "addon_type": "leather", "leather_key": "BUFFALO"}
        ],
        [
            {"col_id": "col_desc", "value": "LEATHER", "style_context": "footer_addon", "addon_type": "leather", "leather_key": "COW"},
            {"col_id": "col_pallet_count", "value": "{cow_pallet_count}", "style_context": "footer_addon", "addon_type": "leather", "leather_key": "COW", "is_pallet": True},
            {"col_id": "col_qty", "value": "{cow_col_qty}", "style_context": "footer_addon", "addon_type": "leather", "leather_key": "COW"}
        ]
    ]
    footer_config = FooterConfigModel(rows=rows)
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        footer_config=footer_config,
        sheet_name="Packing list"
    )
    builder.build()
    
    # Assert Buffalo row (row 0)
    assert grid._grid[0][1].value == "BUFFALO LEATHER"
    assert grid._grid[0][2].value == 2
    assert grid._grid[0][3].value == 100
    
    # Assert Cow row (row 1)
    assert grid._grid[1][1].value == "LEATHER"
    assert grid._grid[1][2].value == 3
    assert grid._grid[1][3].value == 150

def test_declarative_rows_schema():
    column_mapping = {"col_desc": 1, "col_pallet_count": 2, "col_qty": 3}
    style_registry = StyleRegistry({"columns": {}, "row_contexts": {}})
    grid = Grid(column_mapping, style_registry)
    grid.set_section_bounds("data", 5, 10)  # mock data section bounds
    
    declarative_rows = [
        [
            {"col_id": "col_desc", "value": "VAT (7%):", "style_context": "footer"},
            {"col_id": "col_qty", "formula": "SUM", "target_section": "data", "style_context": "footer"},
            {"col_id": "col_pallet_count", "value": "{pallet_count} PALLETS", "style_context": "footer"}
        ]
    ]
    
    footer_config = FooterConfigModel(rows=declarative_rows)
    footer_data = FooterData(
        footer_row_start_idx=11, data_start_row=5, data_end_row=10, total_pallets=8,
        weight_summary=None, leather_summary=None
    )
    
    builder = TableFooterBuilder(
        grid=grid, footer_data=footer_data,
        footer_config=footer_config,
        pallet_count=8
    )
    
    builder.build()
    
    # Assert written value and formatted text
    assert grid._grid[0][1].value == "VAT (7%):"
    assert grid._grid[0][2].value == "8 PALLETS"
    # Assert formula resolved: column 3 is C
    assert grid._grid[0][3].value == "=SUM(C5:C10)"
