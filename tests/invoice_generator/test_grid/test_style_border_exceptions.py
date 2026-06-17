import pytest
import json
from pathlib import Path
from core.invoice_generator.config.config_loader import BundledConfigLoader
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_border_exceptions_resolved_correctly(tmp_path):
    # Prepare dummy bundle config with defaults.borders.exceptions
    config_data = {
        "_meta": {"config_version": "2.2_strict_mode"},
        "processing": {"sheets": ["Invoice"], "data_sources": {"Invoice": "aggregation"}},
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
                    "data": {
                        "bold": False,
                        "font_size": 12,
                        "font_name": "Arial",
                        "border_style": "thin"
                    }
                }
            }
        },
        "layout_bundle": {},
        "data_bundle": {}
    }
    
    # Write config file
    config_file = tmp_path / "test_config.json"
    with open(config_file, "w") as f:
        json.dump(config_data, f)
        
    loader = BundledConfigLoader(config_file)
    styling_config = loader.get_styling_config("Invoice")
    
    # Assert exception merged into columns
    assert styling_config["columns"]["col_static"]["border_style"] == "sides_only"
    assert "border_style" not in styling_config["columns"]["col_qty"]
    
    # Assert StyleRegistry uses border_style
    registry = StyleRegistry(styling_config)
    style = registry.get_style("col_static", context="data")
    assert style["border_style"] == "sides_only"
    
    style_qty = registry.get_style("col_qty", context="data")
    assert style_qty["border_style"] == "thin"
