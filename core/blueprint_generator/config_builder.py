"""
Bundle Builder - Generates JF-style bundle config from template analysis.

This module takes TemplateAnalysisResult and builds a complete bundle config
in the format expected by the invoice generator.
"""

import logging
from datetime import datetime
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

from .excel_scanner import TemplateAnalysisResult, SheetAnalysis, ColumnInfo
from .rules import BlueprintRules
from core.utils.snitch import snitch

logger = logging.getLogger(__name__)


class ConfigBuilder:
    """Builds bundle config from template analysis."""
    
    # Default field mappings for data_flow (IDENTITY MAPPING - Keys match Column IDs)
    # MINIMALIST: No source_key/value, no redundant column mapping.
    # We rely on data_preparer's smart lookup and auto-mapping.
    FIELD_MAPPINGS = {
        # Aggregation sheets (Invoice, Contract)
        "aggregation": {
            "col_po": {},
            "col_item": {},
            "col_unit_price": {},
            "col_desc": {"fallback_on_none": "LEATHER", "fallback_on_DAF": "LEATHER"},
            "col_qty_sf": {},
            "col_amount": {"formula": "{col_qty_sf} * {col_unit_price}"},
        },
        # Processed tables (Packing list)
        "processed_tables_multi": {
            "col_po": {},
            "col_item": {},
            "col_desc": {"fallback_on_none": "LEATHER", "fallback_on_DAF": "LEATHER"},
            "col_qty_pcs": {},
            "col_qty_sf": {},
            "col_net": {},
            "col_gross": {},
            "col_cbm": {},
        }
    }
    
    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)
    
    @snitch
    def build_config(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """
        Build complete bundle config from template analysis.
        Matches the confirmed Master Config structure (v2.1).
        """
        self.logger.info(f"Building blueprint config for {analysis.customer_code}")
        
        bundle = {
            "_meta": self._build_meta(analysis),
            "data_preparation_module_hint": self._build_data_prep_hints(),
            "features": self._build_features(),
            "extensions": self._build_extensions(),
            "processing": self._build_processing(analysis),
            "styling_bundle": self._build_styling_bundle(analysis),
            "layout_bundle": self._build_layout_bundle(analysis),
            "defaults": self._build_defaults()
        }
        
        return bundle
    
    def _build_meta(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build _meta section."""
        return {
            "config_version": "2.2_strict_mode",
            "customer": analysis.customer_code,
            "created_at": datetime.now().strftime("%Y-%m-%d"),
            "compatibility": "strict",
            "description": f"Auto-generated Master Config for {analysis.customer_code}",
            "source_template": analysis.file_path,
            "generator": "blueprint_generator"
        }
    
    def _build_data_prep_hints(self) -> Dict[str, Any]:
        """Build data preparation hints section."""
        return {
            "priority": ["po"],
            "numbers_per_group_by_po": 7
        }
    
    def _build_features(self) -> Dict[str, Any]:
        """Build features section."""
        return {
            "_comment": "Feature flags for optional/experimental functionality",
            "enable_text_replacement": True,
            "enable_conditional_formatting": False,
            "enable_data_validation": False,
            "enable_auto_calculations": True,
            "enable_print_area": False,
            "enable_page_breaks": False,
            "debug_mode": False
        }
    
    def _build_extensions(self) -> Dict[str, Any]:
        """Build extensions section."""
        return {
            "_comment": "Custom extensions - add new features here",
            "_available_hooks": ["pre_build", "post_build", "pre_style", "post_style"],
            "custom": {}
        }
    
    def _build_processing(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build processing section."""
        sheets = [sheet.name for sheet in analysis.sheets]
        data_sources = {sheet.name: sheet.data_source for sheet in analysis.sheets}
        
        return {
            "sheets": sheets,
            "data_sources": data_sources,
            "source_file": analysis.file_path
        }

    def _build_styling_bundle(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build styling_bundle section."""
        styling = {
            "_comment": "Centralized styling - explicit per sheet",
            "defaults": {
                "borders": {
                    "_comment": "Border configuration",
                    "default_border": "full_grid",
                    "default_style": "thin",
                    "exceptions": {
                        "col_static": "side_only"
                    }
                }
            }
        }
        
        for sheet in analysis.sheets:
            styling[sheet.name] = self._build_sheet_styling(sheet)
        
        return styling
    
    def _build_sheet_styling(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build styling for a single sheet."""
        sheet_styling = {
            "_comment": "ID-driven styling with columns + row_contexts",
            "columns": {},
            "row_contexts": {}
        }
        
        # Build column styles
        for col in sheet.columns:
            col_style = {
                "format": col.format,
                "alignment": col.alignment,
                "width": round(col.width, 2)
            }
            if col.wrap_text:
                col_style["wrap_text"] = True
            
            sheet_styling["columns"][col.id] = col_style
            
            # Add child column styles
            for child in col.children:
                child_style = {
                    "format": child.format,
                    "alignment": child.alignment,
                    "width": round(child.width, 2)
                }
                sheet_styling["columns"][child.id] = child_style
        
        # Build row context styles
        sheet_styling["row_contexts"] = {
            "header": {
                "bold": True,
                "font_size": sheet.header_font.get("size", 12),
                "font_name": sheet.header_font.get("name", "Times New Roman"),
                "border_style": "thin",
                "row_height": sheet.row_heights.get("header", 35)
            },
            "data": {
                "bold": False,
                "font_size": sheet.data_font.get("size", 12),
                "font_name": sheet.data_font.get("name", "Times New Roman"),
                "border_style": "thin",
                "row_height": sheet.row_heights.get("data", 27)
            },
            "footer": {
                "bold": True,
                "font_size": sheet.header_font.get("size", 12),
                "font_name": sheet.header_font.get("name", "Times New Roman"),
                "border_style": "thin",
                "row_height": sheet.row_heights.get("footer", 35)
            }
        }
        
        # Add footer add-ons for invoice sheets
        if sheet.data_source == "aggregation":
            sheet_styling["row_contexts"]["footer"]["add_ons"] = {
                "weight_summary": {
                    "enabled": True,
                    "label_col_id": "col_po",
                    "value_col_id": "col_item",
                    "mode": ["daf", "standard"]
                }
            }
        
        return sheet_styling
    
    def _build_layout_bundle(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build layout_bundle section."""
        layout = {
            "_comment": "Layout configuration - structure, data flow, content, and footer"
        }
        
        for sheet in analysis.sheets:
            layout[sheet.name] = self._build_sheet_layout(sheet)
        
        return layout
    
    def _build_sheet_layout(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build layout for a single sheet."""
        layout = {
            "_sections": ["structure", "data_flow", "content", "footer"],
            "structure": self._build_structure(sheet),
            "data_flow": self._build_data_flow(sheet),
            "content": self._build_content(sheet),
            "footer": self._build_footer(sheet)
        }
        
        return layout
    
    def _build_structure(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build structure section for a sheet."""
        columns = []
        
        for col in sheet.columns:
            col_def = {
                "id": col.id,
                "header": col.header
            }
            
            if col.format != "@":
                col_def["format"] = col.format
            
            if col.rowspan > 1:
                col_def["rowspan"] = col.rowspan
            
            if col.colspan > 1:
                col_def["colspan"] = col.colspan
            
            # Add children if present
            if col.children:
                col_def["children"] = []
                for child in col.children:
                    child_def = {
                        "id": child.id,
                        "header": child.header
                    }
                    if child.format != "@":
                        child_def["format"] = child.format
                    col_def["children"].append(child_def)
            
            columns.append(col_def)
        
        return {
            "header_row": sheet.header_row,
            "columns": columns
        }
    
    def _build_data_flow(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build data_flow section for a sheet."""
        mappings = {}
        
        # Get default mappings for this data source type
        default_mappings = self.FIELD_MAPPINGS.get(sheet.data_source, {})
        
        # Build mapping for each column
        all_columns = sheet.columns.copy()
        for col in sheet.columns:
            all_columns.extend(col.children)
        
        self.logger.info(f"  Mapping data flow for sheet '{sheet.name}' ({len(sheet.columns)} cols)...")

        for col in all_columns:
            # IDENTITY MAPPING implies field_name == col.id
            field_name = col.id
            
            if col.id in default_mappings:
                # Start with empty mapping or whatever is in default
                mapping = default_mappings[col.id].copy()
                
                # [Smart Feature] Inject Dynamic Description Fallback
                if col.id == "col_desc" and sheet.static_content_hints:
                     dynamic_fallback = sheet.static_content_hints.get("description_fallback")
                     if dynamic_fallback:
                         mapping["fallback_on_none"] = dynamic_fallback
                         mapping["fallback_on_DAF"] = dynamic_fallback
                         self.logger.info(f"    [Smart]    Updated col_desc fallback to '{dynamic_fallback}'")
                
                # ONLY add "column" if it's different from the field key (which it isn't here)
                mappings[field_name] = mapping
                self.logger.info(f"    [Explicit] Mapped {col.id} -> {field_name} (Minimal)")

            elif col.id not in ["col_static", "col_qty_header", "col_no"]:
                # Create basic mapping for unknown columns
                mappings[field_name] = {} 
                self.logger.info(f"    [Auto]     Mapped {col.id} -> {field_name} (Minimal)")
            else:
                 self.logger.debug(f"    [Skipped]  {col.id} (Structural/Static)")
            
            # [CRITICAL FIX] Inject source_key for aggregation sheets
            # Data Parser reads rows as lists, so we need the index to find the value.
            if sheet.data_source in ["aggregation", "DAF_aggregation"] and field_name in mappings:
                # Excel 1-based -> Python 0-based
                mappings[field_name]["source_key"] = col.col_index - 1
                self.logger.info(f"    [Explicit] Injected source_key={mappings[field_name]['source_key']} for {field_name}")
        
        return {"mappings": mappings}
    
    def _col_id_to_field_name(self, col_id: str) -> str:
        """Convert column ID to field name."""
        # IDENTITY MAPPING: Return the column ID directly.
        # This ensures that the generated config uses the same keys as the data parser output.
        return col_id
    
    def _build_content(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build content section for a sheet."""
        content = {"static": {}}
        
        # Add static content hints
        if sheet.static_content_hints:
            content["static"] = sheet.static_content_hints
        elif "col_static" in [c.id for c in sheet.columns]:
            # Default static content
            content["static"]["col_static"] = [
                "VENDOR#:",
                "Des: LEATHER",
                "MADE IN CAMBODIA"
            ]
        
        return content
    
    def _build_footer(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build footer section for a sheet."""
        # Find appropriate columns for footer
        col_ids = [c.id for c in sheet.columns]
        for c in sheet.columns:
            col_ids.extend([child.id for child in c.children])
        
        # Determine total text column
        total_col = "col_po"
        if "col_no" in col_ids:
            total_col = "col_no"
        
        # Determine pallet count column  
        pallet_col = "col_desc" if "col_desc" in col_ids else "col_item"
        
        # Determine sum columns
        sum_cols = []
        default_sum = BlueprintRules.DEFAULT_FOOTER_SUMS.get(sheet.data_source, [])
        for col_id in default_sum:
            if col_id in col_ids:
                sum_cols.append(col_id)
        
        footer = {
            "total_text_column_id": total_col,
            "total_text": "TOTAL OF:",
            "pallet_count_column_id": pallet_col,
            "sum_column_ids": sum_cols,
            "merge_rules": [],
            "add_ons": self._build_footer_addons(sheet)
        }
        
        return footer
    
    def _build_footer_addons(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build footer add-ons configuration."""
        add_ons = {}
        
        # before_footer (HS.CODE line)
        if sheet.data_source == "processed_tables_multi":
            add_ons["before_footer"] = {
                "enabled": True,
                "column_id": "col_po",
                "text": "LEATHER (HS.CODE: 4107.12.00)",
                "merge": 2
            }
        else:
            add_ons["before_footer"] = {
                "enabled": False,
                "column_id": "col_po",
                "text": "HS.CODE: 4107.12.00"
            }
        
        # weight_summary
        add_ons["weight_summary"] = {
            "enabled": sheet.data_source == "aggregation",
            "label_col_id": "col_po",
            "value_col_id": "col_item",
            "mode": ["daf", "standard"]
        }
        
        # leather_summary (for packing list)
        add_ons["leather_summary"] = {
            "enabled": sheet.data_source == "processed_tables_multi",
            "mode": ["daf", "standard"]
        }
        
        return add_ons
    
    def _build_defaults(self) -> Dict[str, Any]:
        """Build defaults section."""
        return {
            "footer": {
                "show_total": True,
                "show_pallet_count": True,
                "total_text": "TOTAL:",
                "merge_total_cells": True,
                "sum_columns": ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm"]
            }
        }


if __name__ == "__main__":
    import sys
    import json
    from .excel_scanner import ExcelLayoutScanner
    
    logging.basicConfig(level=logging.INFO)
    
    if len(sys.argv) < 2:
        print("Usage: python config_builder.py <template.xlsx>")
        sys.exit(1)
    
    scanner = ExcelLayoutScanner()
    result = scanner.scan_template(sys.argv[1])
    
    builder = ConfigBuilder()
    bundle = builder.build_config(result)
    
    print(json.dumps(bundle, indent=2, ensure_ascii=False))
