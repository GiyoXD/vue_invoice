"""
Config Converter - Converts old-format configs to new bundle format.

This module converts existing configs (like CLW_config.json) to the 
JF bundle format used by the invoice generator.
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional
from datetime import datetime

logger = logging.getLogger(__name__)


class LegacyConfigMigrator:
    """Converts old config format to new bundle format."""
    
    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)
    
    def convert(self, old_config_path: str, customer_code: str = None) -> Dict[str, Any]:
        """
        Convert old config to bundle format.
        
        Args:
            old_config_path: Path to old config JSON file
            customer_code: Optional customer code (derived from filename if not provided)
            
        Returns:
            Bundle config dictionary
        """
        path = Path(old_config_path)
        
        if not customer_code:
            # Extract from filename like "CLW_config.json" -> "CLW"
            customer_code = path.stem.replace("_config", "").upper()
        
        self.logger.info(f"Converting config for {customer_code}")
        
        with open(path, 'r', encoding='utf-8') as f:
            old_config = json.load(f)
        
        bundle = {
            "_meta": self._build_meta(customer_code, old_config_path),
            "data_preparation_module_hint": {"priority": ["po"], "numbers_per_group_by_po": 7},
            "features": self._build_features(),
            "extensions": self._build_extensions(),
            "processing": self._build_processing(old_config),
            "styling_bundle": self._build_styling_bundle(old_config),
            "layout_bundle": self._build_layout_bundle(old_config),
            "defaults": self._build_defaults()
        }
        
        return bundle
    
    def _build_meta(self, customer_code: str, source_path: str) -> Dict[str, Any]:
        return {
            "config_version": "2.1_converted",
            "customer": customer_code,
            "created": datetime.now().strftime("%Y-%m-%d"),
            "description": f"Converted from old config for {customer_code}",
            "source_config": source_path,
            "generator": "config_converter"
        }
    
    def _build_features(self) -> Dict[str, Any]:
        return {
            "_comment": "Feature flags",
            "enable_text_replacement": False,
            "enable_conditional_formatting": False,
            "enable_data_validation": False,
            "enable_auto_calculations": True,
            "enable_print_area": False,
            "enable_page_breaks": False,
            "debug_mode": False
        }
    
    def _build_extensions(self) -> Dict[str, Any]:
        return {
            "_comment": "Custom extensions",
            "_available_hooks": ["pre_build", "post_build", "pre_style", "post_style"],
            "custom": {}
        }
    
    def _build_processing(self, old_config: Dict) -> Dict[str, Any]:
        sheets = old_config.get("sheets_to_process", [])
        
        # Convert data source names
        data_sources = {}
        sheet_map = old_config.get("sheet_data_map", {})
        for sheet, source in sheet_map.items():
            if source == "processed_tables_data":
                data_sources[sheet] = "processed_tables_multi"
            else:
                data_sources[sheet] = source
        
        return {
            "sheets": sheets,
            "data_sources": data_sources
        }
    
    def _build_styling_bundle(self, old_config: Dict) -> Dict[str, Any]:
        styling = {
            "_comment": "Centralized styling",
            "defaults": {
                "borders": {
                    "default_border": "full_grid",
                    "default_style": "thin",
                    "exceptions": {"col_static": "side_only"}
                }
            }
        }
        
        data_mapping = old_config.get("data_mapping", {})
        
        for sheet_name, sheet_data in data_mapping.items():
            styling[sheet_name] = self._build_sheet_styling(sheet_name, sheet_data)
        
        return styling
    
    def _build_sheet_styling(self, sheet_name: str, sheet_data: Dict) -> Dict[str, Any]:
        columns = {}
        
        # Extract column styles from headers and mappings
        headers = sheet_data.get("header_to_write", [])
        mappings = sheet_data.get("mappings", {})
        styling_data = sheet_data.get("styling", {})
        
        for header in headers:
            col_id = header.get("id")
            if not col_id:
                continue
            
            col_style = {
                "format": "@",
                "alignment": "center",
                "width": 15
            }
            
            # Check for number format in mappings
            for field, mapping in mappings.items():
                if mapping.get("id") == col_id:
                    if "number_format" in mapping:
                        col_style["format"] = mapping["number_format"]
                    break
            
            # Check column widths from styling
            col_widths = styling_data.get("column_widths", {})
            if col_id in col_widths:
                col_style["width"] = col_widths[col_id]
            
            columns[col_id] = col_style
        
        # Extract font info
        header_font = styling_data.get("header_font", {})
        data_font = styling_data.get("data_font", {})
        
        row_contexts = {
            "header": {
                "bold": True,
                "font_size": header_font.get("size", 12),
                "font_name": header_font.get("name", "Times New Roman"),
                "border_style": "thin",
                "row_height": 35
            },
            "data": {
                "bold": False,
                "font_size": data_font.get("size", 12),
                "font_name": data_font.get("name", "Times New Roman"),
                "border_style": "thin",
                "row_height": 27
            },
            "footer": {
                "bold": True,
                "font_size": header_font.get("size", 12),
                "font_name": header_font.get("name", "Times New Roman"),
                "border_style": "thin",
                "row_height": 35
            }
        }
        
        return {
            "_comment": f"Styling for {sheet_name}",
            "columns": columns,
            "row_contexts": row_contexts
        }
    
    def _build_layout_bundle(self, old_config: Dict) -> Dict[str, Any]:
        layout = {"_comment": "Layout configuration"}
        
        data_mapping = old_config.get("data_mapping", {})
        sheet_map = old_config.get("sheet_data_map", {})
        
        for sheet_name, sheet_data in data_mapping.items():
            data_source = sheet_map.get(sheet_name, "aggregation")
            layout[sheet_name] = self._build_sheet_layout(sheet_name, sheet_data, data_source)
        
        return layout
    
    def _build_sheet_layout(self, sheet_name: str, sheet_data: Dict, data_source: str) -> Dict[str, Any]:
        return {
            "_sections": ["structure", "data_flow", "content", "footer"],
            "structure": self._build_structure(sheet_data),
            "data_flow": self._build_data_flow(sheet_data, data_source),
            "content": self._build_content(sheet_data),
            "footer": self._build_footer(sheet_data, data_source)
        }
    
    def _build_structure(self, sheet_data: Dict) -> Dict[str, Any]:
        header_row = sheet_data.get("start_row", 20) - 1  # Convert to 0-based then back
        headers = sheet_data.get("header_to_write", [])
        
        columns = []
        for header in headers:
            col_def = {
                "id": header.get("id", ""),
                "header": header.get("text", "")
            }
            
            if header.get("rowspan", 1) > 1:
                col_def["rowspan"] = header["rowspan"]
            if header.get("colspan", 1) > 1:
                col_def["colspan"] = header["colspan"]
            
            # Add child columns if this is a parent header
            children = header.get("children", [])
            if children:
                col_def["children"] = [
                    {"id": c.get("id"), "header": c.get("text")}
                    for c in children
                ]
            
            columns.append(col_def)
        
        return {
            "header_row": sheet_data.get("start_row", 20),
            "columns": columns
        }
    
    def _build_data_flow(self, sheet_data: Dict, data_source: str) -> Dict[str, Any]:
        mappings = {}
        old_mappings = sheet_data.get("mappings", {})
        
        for field_name, mapping in old_mappings.items():
            new_mapping = {"column": mapping.get("id", "")}
            
            # Handle source key/value
            if "key_index" in mapping:
                new_mapping["source_key"] = mapping["key_index"]
            if "value_key" in mapping:
                new_mapping["source_value"] = mapping["value_key"]
            
            # Handle fallbacks
            if "fallback_on_none" in mapping:
                new_mapping["fallback_on_none"] = mapping["fallback_on_none"]
            if "fallback_on_DAF" in mapping:
                new_mapping["fallback_on_DAF"] = mapping["fallback_on_DAF"]
            
            # Handle formulas
            if "formula" in mapping:
                new_mapping["formula"] = mapping["formula"]
            
            mappings[field_name] = new_mapping
        
        return {"mappings": mappings}
    
    def _build_content(self, sheet_data: Dict) -> Dict[str, Any]:
        static_content = sheet_data.get("static_content", {})
        
        # Default static content for Mark column
        if not static_content:
            static_content = {
                "col_static": [
                    "VENDOR#:",
                    "Des: LEATHER",
                    "MADE IN CAMBODIA"
                ]
            }
        
        return {"static": static_content}
    
    def _build_footer(self, sheet_data: Dict, data_source: str) -> Dict[str, Any]:
        footer_config = sheet_data.get("footer_config", {})
        
        # Default footer configuration
        footer = {
            "total_text_column_id": footer_config.get("total_text_column_id", "col_po"),
            "total_text": footer_config.get("total_text", "TOTAL OF:"),
            "pallet_count_column_id": footer_config.get("pallet_count_column_id", "col_desc"),
            "sum_column_ids": footer_config.get("sum_column_ids", []),
            "merge_rules": [],
            "add_ons": {}
        }
        
        # Set default sum columns based on data source
        if not footer["sum_column_ids"]:
            if data_source == "processed_tables_data":
                footer["sum_column_ids"] = ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm"]
            else:
                footer["sum_column_ids"] = ["col_qty_sf", "col_amount"]
        
        # Add-ons
        static_before = sheet_data.get("static_content_before_footer", {})
        footer["add_ons"]["before_footer"] = {
            "enabled": bool(static_before),
            "column_id": "col_po",
            "text": list(static_before.values())[0] if static_before else "HS.CODE: 4107.12.00"
        }
        
        footer["add_ons"]["weight_summary"] = {
            "enabled": data_source == "aggregation",
            "label_col_id": "col_po",
            "value_col_id": "col_item",
            "mode": ["daf", "standard"]
        }
        
        footer["add_ons"]["leather_summary"] = {
            "enabled": data_source == "processed_tables_data",
            "mode": ["daf", "standard"]
        }
        
        return footer
    
    def _build_defaults(self) -> Dict[str, Any]:
        return {
            "footer": {
                "show_total": True,
                "show_pallet_count": True,
                "total_text": "TOTAL:",
                "merge_total_cells": True,
                "sum_columns": ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm"]
            }
        }


def main():
    """CLI for converting old configs."""
    import argparse
    import sys
    
    parser = argparse.ArgumentParser(description="Convert old config to bundle format")
    parser.add_argument('config_path', help='Path to old config JSON file')
    parser.add_argument('-o', '--output', help='Output path (default: stdout)')
    parser.add_argument('-c', '--customer', help='Customer code (derived from filename if not provided)')
    
    args = parser.parse_args()
    
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    
    converter = LegacyConfigMigrator()
    bundle = converter.convert(args.config_path, args.customer)
    
    output = json.dumps(bundle, indent=2, ensure_ascii=False)
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"Saved to {args.output}")
    else:
        print(output)


if __name__ == "__main__":
    main()
