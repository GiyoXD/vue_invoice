import pytest
import json
from pathlib import Path
from core.invoice_generator.config.config_reader import ConfigFileReader
from core.invoice_generator.config.config_store import ConfigStore
from core.invoice_generator.styling.style_registry import StyleRegistry
from core.invoice_generator.styling.border_resolver import BorderResolver, BORDER_PATTERNS
from core.invoice_generator.styling.dimension_registry import DimensionRegistry
from core.invoice_generator.builders.table.table_grid import TableGrid
from core.models.cell import BorderStyle


def _make_grid_with_data(styling_config, rows=3):
    """Helper: creates a TableGrid with header + data + footer sections populated."""
    registry = StyleRegistry(styling_config)
    column_mapping = {"col_desc": 1, "col_qty": 2, "col_price": 3}
    dim_registry = DimensionRegistry({})
    
    grid = TableGrid(
        column_mapping=column_mapping,
        style_registry=registry,
        dimension_registry=dim_registry
    )
    grid.set_start_row(5)
    
    # Header section (2 rows)
    grid.mark_section_start("header")
    for col_id in column_mapping:
        grid.write(0, col_id, f"Header-{col_id}", context="header")
        grid.write(1, col_id, f"Sub-{col_id}", context="header")
    grid.advance_row(2)
    grid.mark_section_end("header")
    
    # Data section
    grid.mark_section_start("data")
    for i in range(rows):
        for col_id in column_mapping:
            grid.write(i, col_id, f"data-{col_id}-{i}", context="data")
    grid.advance_row(rows)
    grid.mark_section_end("data")
    
    # Footer section (1 row)
    grid.mark_section_start("footer")
    for col_id in column_mapping:
        grid.write(0, col_id, f"footer-{col_id}", context="footer")
    grid.advance_row(1)
    grid.mark_section_end("footer")
    
    return grid, column_mapping


def test_full_grid_all_cells_get_full_border(tmp_path):
    """full_grid mode: every cell gets thin borders on all 4 sides."""
    config_data = {
        "_meta": {"config_version": "2.2_strict_mode"},
        "processing": {"Invoice": "aggregation"},
        "styling_bundle": {
            "Invoice": {
                "default_border": "full_grid",
                "columns": {
                    "col_desc": {"format": "@", "alignment": "left"},
                    "col_qty": {"format": "#,##0", "alignment": "center"},
                    "col_price": {"format": "#,##0.00", "alignment": "right"}
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
    
    config_file = tmp_path / "test_full_grid.json"
    with open(config_file, "w") as f:
        json.dump(config_data, f)
        
    config_data, template_data = ConfigFileReader.load(config_file)
    loader = ConfigStore(config_data, template_data)
    styling_config = loader.get_styling_config("Invoice")
    
    grid, column_mapping = _make_grid_with_data(styling_config)
    
    resolver = BorderResolver(default_border="full_grid")
    resolver.apply(grid)
    
    # Every cell should have full thin border
    for row_idx in sorted(grid._grid.keys()):
        for col_idx in grid._grid[row_idx]:
            cell = grid._grid[row_idx][col_idx]
            assert cell.style is not None, f"Cell ({row_idx},{col_idx}) has no style"
            assert cell.style.border is not None, f"Cell ({row_idx},{col_idx}) has no border"
            assert cell.style.border.left == "thin"
            assert cell.style.border.right == "thin"
            assert cell.style.border.top == "thin"
            assert cell.style.border.bottom == "thin"


def test_side_grid_data_rows_get_sides_only(tmp_path):
    """side_grid mode: data rows get sides_only, header/footer get full border."""
    config_data = {
        "_meta": {"config_version": "2.2_strict_mode"},
        "processing": {"Invoice": "aggregation"},
        "styling_bundle": {
            "Invoice": {
                "default_border": "side_grid",
                "columns": {
                    "col_desc": {"format": "@", "alignment": "left"},
                    "col_qty": {"format": "#,##0", "alignment": "center"},
                    "col_price": {"format": "#,##0.00", "alignment": "right"}
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
    
    config_file = tmp_path / "test_side_grid.json"
    with open(config_file, "w") as f:
        json.dump(config_data, f)
        
    config_data, template_data = ConfigFileReader.load(config_file)
    loader = ConfigStore(config_data, template_data)
    styling_config = loader.get_styling_config("Invoice")
    
    grid, column_mapping = _make_grid_with_data(styling_config, rows=3)
    
    resolver = BorderResolver(default_border="side_grid")
    resolver.apply(grid)
    
    all_rows = sorted(grid._grid.keys())
    last_row = all_rows[-1]
    
    # Header rows (first 2) should have full border
    for row_idx in all_rows[:2]:
        for col_idx in grid._grid[row_idx]:
            cell = grid._grid[row_idx][col_idx]
            assert cell.style.border.left == "thin"
            assert cell.style.border.right == "thin"
            assert cell.style.border.top == "thin"
            assert cell.style.border.bottom == "thin"
    
    # Data rows (middle) should have sides_only
    for row_idx in all_rows[2:-1]:
        for col_idx in grid._grid[row_idx]:
            cell = grid._grid[row_idx][col_idx]
            assert cell.style.border.left == "thin"
            assert cell.style.border.right == "thin"
            assert cell.style.border.top is None
            assert cell.style.border.bottom is None

    # Last row (footer) should have bottom border to close table
    for col_idx in grid._grid[last_row]:
        cell = grid._grid[last_row][col_idx]
        assert cell.style.border.left == "thin"
        assert cell.style.border.right == "thin"
        assert cell.style.border.bottom == "thin", f"Last row cell col={col_idx} missing bottom border"


def test_column_override_preserved_in_side_grid(tmp_path):
    """Column with explicit border_style override is not affected by side_grid mode."""
    config_data = {
        "_meta": {"config_version": "2.2_strict_mode"},
        "processing": {"Invoice": "aggregation"},
        "styling_bundle": {
            "Invoice": {
                "default_border": "side_grid",
                "columns": {
                    "col_desc": {"format": "@", "alignment": "left"},
                    "col_qty": {"format": "#,##0", "alignment": "center"},
                    "col_price": {"format": "#,##0.00", "alignment": "right", "border_style": "thin"}
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
    
    config_file = tmp_path / "test_col_override.json"
    with open(config_file, "w") as f:
        json.dump(config_data, f)
        
    config_data, template_data = ConfigFileReader.load(config_file)
    loader = ConfigStore(config_data, template_data)
    styling_config = loader.get_styling_config("Invoice")
    
    grid, column_mapping = _make_grid_with_data(styling_config, rows=3)
    
    # col_price has explicit "thin" override
    resolver = BorderResolver(
        default_border="side_grid",
        column_overrides={"col_price": "thin"}
    )
    resolver.apply(grid)
    
    all_rows = sorted(grid._grid.keys())
    col_price_idx = column_mapping["col_price"]
    col_desc_idx = column_mapping["col_desc"]
    
    # Data rows: col_price should have full thin (override), col_desc should have sides_only
    for row_idx in all_rows[2:-1]:  # skip header rows and last row
        price_cell = grid._grid[row_idx][col_price_idx]
        assert price_cell.style.border.top == "thin", f"col_price override not applied at row {row_idx}"
        assert price_cell.style.border.bottom == "thin"
        
        desc_cell = grid._grid[row_idx][col_desc_idx]
        assert desc_cell.style.border.top is None, f"col_desc should be sides_only at row {row_idx}"
        assert desc_cell.style.border.bottom is None


def test_last_row_always_has_bottom_border():
    """The absolute last row of the grid always gets a bottom border, regardless of pattern."""
    resolver = BorderResolver(default_border="side_grid")
    
    # Create a minimal grid with just data (no header/footer)
    from core.invoice_generator.styling.style_registry import StyleRegistry
    styling_config = {
        "columns": {"col_a": {"format": "@", "alignment": "center"}},
        "row_contexts": {"data": {"bold": False, "font_size": 12, "font_name": "Arial"}}
    }
    registry = StyleRegistry(styling_config)
    grid = TableGrid(
        column_mapping={"col_a": 1},
        style_registry=registry
    )
    grid.set_start_row(1)
    grid.mark_section_start("data")
    grid.write(0, "col_a", "row1", context="data")
    grid.write(1, "col_a", "row2", context="data")
    grid.write(2, "col_a", "row3", context="data")
    grid.advance_row(3)
    grid.mark_section_end("data")
    
    resolver.apply(grid)
    
    all_rows = sorted(grid._grid.keys())
    last_row = all_rows[-1]
    
    # Last row must have bottom border
    cell = grid._grid[last_row][1]
    assert cell.style.border.bottom == "thin"
    
    # Non-last data rows should NOT have bottom (sides_only)
    for row_idx in all_rows[:-1]:
        cell = grid._grid[row_idx][1]
        assert cell.style.border.bottom is None, f"Row {row_idx} should not have bottom border"


def test_border_patterns_are_correct():
    """Verify all named border patterns produce the expected BorderStyle."""
    thin = BORDER_PATTERNS["thin"]
    assert thin.left == "thin" and thin.right == "thin" and thin.top == "thin" and thin.bottom == "thin"
    
    sides = BORDER_PATTERNS["sides_only"]
    assert sides.left == "thin" and sides.right == "thin" and sides.top is None and sides.bottom is None
    
    no_top = BORDER_PATTERNS["no_top"]
    assert no_top.left == "thin" and no_top.right == "thin" and no_top.top is None and no_top.bottom == "thin"
    
    no_bottom = BORDER_PATTERNS["no_bottom"]
    assert no_bottom.left == "thin" and no_bottom.right == "thin" and no_bottom.top == "thin" and no_bottom.bottom is None
    
    assert BORDER_PATTERNS["none"] is None


def test_style_registry_no_longer_returns_border():
    """StyleRegistry.get_style() should NOT return border_style in its output."""
    styling_config = {
        "columns": {"col_a": {"format": "@", "alignment": "center", "border_style": "thin"}},
        "row_contexts": {"data": {"bold": False, "font_size": 12, "font_name": "Arial"}}
    }
    registry = StyleRegistry(styling_config)
    style = registry.get_style("col_a", context="data")
    
    # border_style should NOT be in the merged style
    assert "border_style" not in style
