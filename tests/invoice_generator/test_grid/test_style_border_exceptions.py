import pytest
import json
from pathlib import Path
from core.invoice_generator.config import ConfigFileReader, ConfigStore
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.styling.border_resolver import BorderResolver
from core.invoice_generator.builders.table.table_grid import TableGrid
from core.models.cell import BorderStyle

def test_border_exceptions_resolved_correctly(tmp_path):
    """Border exceptions from global defaults are merged into column configs
    and correctly used by BorderResolver (not StyleRegistry)."""
    config_data = {
        "_meta": {"config_version": "2.2_strict_mode"},
        "processing": {"Invoice": "aggregation"},
        "styling_bundle": {
            "defaults": {
                "borders": {
                    "default_border": "full_grid",
                    "default_style": "thin",
                    "exceptions": {
                        "col_static": "side_only"
                    }
                }
            },
            "Invoice": {
                "columns": {
                    "col_static": {"format": "@", "alignment": "center"},
                    "col_qty": {"format": "#,##0", "alignment": "center"}
                },
                "row_contexts": {
                    "header": {"bold": True, "font_size": 12, "font_name": "Arial"},
                    "data": {"bold": False, "font_size": 12, "font_name": "Arial"},
                    "footer": {"bold": True, "font_size": 12, "font_name": "Arial"}
                }
            }
        },
        "layout_bundle": {},
        "data_bundle": {}
    }
    
    config_file = tmp_path / "test_config.json"
    with open(config_file, "w") as f:
        json.dump(config_data, f)
        
    config_data, template_data = ConfigFileReader.load(config_file)
    loader = ConfigStore(config_data, template_data)
    styling_config = loader.get_styling_config("Invoice")
    
    # 1. Config store still merges exceptions into column configs
    assert styling_config["columns"]["col_static"]["border_style"] == "sides_only"
    assert "border_style" not in styling_config["columns"]["col_qty"]
    
    # 2. StyleRegistry no longer returns border_style
    registry = StyleRegistry(styling_config)
    style = registry.get_style("col_static", context="data")
    assert "border_style" not in style
    
    # 3. BorderResolver correctly applies the column override
    grid = TableGrid(column_mapping={"col_static": 1, "col_qty": 2}, style_registry=registry)
    grid.set_start_row(1)
    grid.mark_section_start("data")
    grid.write(0, "col_static", "static_val", context="data")
    grid.write(0, "col_qty", 100, context="data")
    grid.advance_row(1)
    grid.mark_section_end("data")
    
    resolver = BorderResolver(
        default_border="full_grid",
        column_overrides={"col_static": "sides_only"}
    )
    resolver.apply(grid)
    
    # col_static should have sides_only (but with bottom on last row)
    static_cell = grid._grid[0][1]
    assert static_cell.style.border.left == "thin"
    assert static_cell.style.border.right == "thin"
    # Last row gets bottom added
    assert static_cell.style.border.bottom == "thin"
    
    # col_qty should have full thin border (no override, full_grid default)
    qty_cell = grid._grid[0][2]
    assert qty_cell.style.border.left == "thin"
    assert qty_cell.style.border.right == "thin"
    assert qty_cell.style.border.top == "thin"
    assert qty_cell.style.border.bottom == "thin"
