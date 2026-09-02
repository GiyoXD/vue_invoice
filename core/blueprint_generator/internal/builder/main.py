"""
Bundle Builder - Generates JF-style bundle config from template analysis.

This module takes TemplateAnalysisResult and builds a complete bundle config
in the format expected by the invoice generator.
"""

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

from ..scanner import TemplateAnalysisResult, SheetAnalysis, ColumnInfo
from ..validator import ConfigValidator, BlueprintLogicValidator
from ..addons import AddonRegistry
from .footer import build_footer
from .summary import build_summary

from core.utils.snitch import snitch
from core.invoice_generator.utils.formula_formatter import extract_decimal_places

logger = logging.getLogger(__name__)


class ConfigBuilder:
    """Builds bundle config from template analysis."""
    
    # Default field mappings by data source type.
    # These define WHICH columns to include in data_flow for each source type.
    # Formulas and fallbacks are NOT here — they come from master_config.json defaults.
    FIELD_MAPPINGS = {
        # Aggregation sheets (Invoice, Contract)
        "aggregation": [
            "col_po", "col_item", "col_desc", "col_qty_sf",
            "col_unit_price", "col_amount", "col_sqm", "col_net", "col_gross", "col_cbm"
        ],
        # Processed tables (Packing list)
        "processed_tables_multi": [
            "col_po", "col_item", "col_desc", "col_qty_pcs",
            "col_qty_sf", "col_net", "col_gross", "col_cbm", "col_sqm"
        ]
    }
    
    def __init__(self, fallback_mappings: Optional[Dict[str, str]] = None):
        self.logger = logging.getLogger(self.__class__.__name__)
        self.fallback_mappings = fallback_mappings or {}
        self._default_mapping_keys = set()
    
    @snitch
    def build_config(self, analysis: TemplateAnalysisResult, fallback_mappings: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """
        Build complete bundle config from template analysis.
        Matches the confirmed Master Config structure (v2.1).
        """
        if fallback_mappings is not None:
            self.fallback_mappings = fallback_mappings
            
        self.logger.info(f"Building blueprint config for {analysis.customer_code}")
        
        bundle = {
            "_meta": self._build_meta(analysis),
            "processing": self._build_processing(analysis),
            "styling_bundle": self._build_styling_bundle(analysis),
            "layout_bundle": self._build_layout_bundle(analysis)
        }
        
        return bundle
    
    def _build_meta(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build _meta section."""
        decimals_map: Dict[str, int] = {}
        try:
            for sheet in analysis.sheets:
                all_cols = []
                for col in sheet.columns:
                    all_cols.append(col)
                    if getattr(col, "children", None):
                        all_cols.extend(col.children)
                
                for col in all_cols:
                    col_id = getattr(col, "id", None)
                    col_fmt = getattr(col, "format", None)
                    if col_id and col_fmt:
                        extracted = extract_decimal_places(col_fmt)
                        if extracted is not None:
                            decimals_map[col_id] = extracted
        except Exception as e:
            self.logger.error(f"Failed to extract column decimals: {e}", exc_info=True)

        return {
            "config_version": "2.2_strict_mode",
            "customer": analysis.customer_code,
            "created_at": datetime.now().strftime("%Y-%m-%d"),
            "compatibility": "strict",
            "description": f"Auto-generated Master Config for {analysis.customer_code}",
            "source_template": analysis.file_path,
            "generator": "blueprint_generator",
            "has_static_sheets": analysis.has_static_sheets,
            "decimals": decimals_map,
            "data_preparation_module_hint": self._build_data_prep_hints()
        }
    
    def _build_data_prep_hints(self) -> Dict[str, Any]:
        """Build data preparation hints section."""
        return {
            "priority": ["po"],
            "numbers_per_group_by_po": 7
        }

    
    def _build_processing(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build processing section."""
        # Map sheet names directly to their data source (processing mode)
        return {sheet.name: sheet.data_source for sheet in analysis.sheets}

    def _build_styling_bundle(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build styling_bundle section."""
        styling = {
            "defaults": {
                "default_border": "full_grid"
            }
        }
        
        for sheet in analysis.sheets:
            styling[sheet.name] = self._build_sheet_styling(sheet)
        
        return styling
    
    def _build_sheet_styling(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build styling for a single sheet."""
        sheet_styling = {
            "columns": {},
            "row_contexts": {}
        }
        
        # Build column styles
        for col in sheet.columns:
            col_style = {
                "format": col.format,
                "alignment": col.alignment
            }
            if col.wrap_text:
                col_style["wrap_text"] = True
            if col.id == 'col_static':
                col_style["border_style"] = "side_only"
            
            sheet_styling["columns"][col.id] = col_style
            
            # Add child column styles
            for child in col.children:
                child_style = {
                    "format": child.format,
                    "alignment": child.alignment
                }
                if child.id == 'col_static':
                    child_style["border_style"] = "side_only"
                sheet_styling["columns"][child.id] = child_style
        
        # Build row context styles
        # Enforce strict style extraction - no guesses.
        h_size = sheet.header_font.get("size")
        h_name = sheet.header_font.get("name")
        d_size = sheet.data_font.get("size")
        d_name = sheet.data_font.get("name")

        if any(v is None for v in [h_size, h_name, d_size, d_name]):
             raise ValueError(f"Sheet '{sheet.name}' is missing mandatory font information. Check template styling.")

        sheet_styling["row_contexts"] = {
            "header": {
                "bold": True,
                "font_size": h_size,
                "font_name": h_name,
                "border_style": "thin",
                "row_height": float(sheet.row_heights.get("header", 1))
            },
            "data": {
                "bold": False,
                "font_size": d_size,
                "font_name": d_name,
                "border_style": "thin",
                "row_height": float(sheet.row_heights.get("data", 1))
            },
            "footer": {
                "bold": True,
                "font_size": h_size,
                "font_name": h_name,
                "border_style": "thin",
                "row_height": float(sheet.row_heights.get("footer", 1))
            },
            "hs_code": {
                "bold": True,
                "font_size": d_size,
                "font_name": d_name,
                "border_style": "thin",
                "row_height": float(sheet.row_heights.get("footer", sheet.row_heights.get("data", 1)))
            },
            "summary": {
                "bold": True,
                "font_size": h_size,
                "font_name": h_name,
                "border_style": "none"
            }


        }
        
        return sheet_styling
    
    def _build_layout_bundle(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """Build layout_bundle section, sourcing defaults from master_config.json."""
        layout = {
            "_comment": "Layout configuration - structure, data flow, content, and footer"
        }

        # Load defaults from master_config.json and cache the column IDs
        master_defaults = self._load_master_defaults()
        if master_defaults:
            layout["defaults"] = master_defaults
            self._default_mapping_keys = set(
                master_defaults.get('data_flow', {}).get('mappings', {}).keys()
            )
            self.logger.info(f"  [Defaults] Loaded {len(self._default_mapping_keys)} default mappings from master_config")
        else:
            self._default_mapping_keys = set()
        
        for sheet in analysis.sheets:
            layout[sheet.name] = self._build_sheet_layout(sheet)
        
        return layout
    
    def _load_master_defaults(self) -> Optional[Dict[str, Any]]:
        """
        Load the defaults section from master_config.json.
        
        Returns:
            The defaults dict from layout_bundle, or None if not found.
        """
        try:
            from core.system_config import sys_config
            master_path = sys_config.blueprints_root / "mapper" / "master_config.json"
            
            if not master_path.exists():
                self.logger.warning(f"Master config not found at {master_path}")
                return None
            
            with open(master_path, 'r', encoding='utf-8') as f:
                master_config = json.load(f)
            
            defaults = master_config.get('layout_bundle', {}).get('defaults')
            if defaults:
                return defaults
            else:
                self.logger.warning("No 'defaults' section found in master_config.json layout_bundle")
                return None
        except Exception as e:
            self.logger.error(f"Failed to load master config defaults: {e}")
            return None
    
    def _build_sheet_layout(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build layout for a single sheet."""
        layout = {
            "_sections": ["structure", "data_flow", "static_payload", "footer", "summary"],
            "structure": self._build_structure(sheet),
            "data_flow": self._build_data_flow(sheet),
            "static_payload": self._build_static_payload(sheet),
            "footer": self._build_footer(sheet),
            "summary": self._build_summary(sheet)
        }

        
        hs_code_info = self._build_hs_code(sheet)
        if hs_code_info:
            layout["_sections"].append("hs_code")
            layout["hs_code"] = hs_code_info
            
        return layout
    
    def _build_structure(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build structure section for a sheet."""
        columns = []
        
        # [Verification] Strict Mode: Ensure all IDs exist in Mapper Config
        # Delegated to BlueprintLogicValidator
        BlueprintLogicValidator.verify_strict_mode(sheet)
        
        for col in sheet.columns:
            col_def = {
                "id": col.id,
                "header": col.header,
                "width": round(col.width, 2)
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
                        "header": child.header,
                        "width": round(child.width, 2)
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
        """
        Build data_flow section for a sheet.
        
        Only emits sheet-specific overrides. Columns covered by 
        layout_bundle.defaults are not duplicated here.
        """
        mappings = {}
        
        # Get known column IDs for this data source type
        known_columns = self.FIELD_MAPPINGS.get(sheet.data_source, [])
        
        # Build mapping for each column (including children)
        all_columns = sheet.columns.copy()
        for col in sheet.columns:
            all_columns.extend(col.children)
        
        self.logger.info(f"  Mapping data flow for sheet '{sheet.name}' ({len(sheet.columns)} cols)...")

        # Load global fallback strategies from DB
        global_fallbacks = {}
        try:
            from core.database.db_manager import get_global_mapping_config
            db_config = get_global_mapping_config()
            if isinstance(db_config, dict):
                global_fallbacks = db_config.get("fallback_strategies", {}) or {}
        except Exception as e:
            self.logger.debug(f"Could not load global fallback strategies from DB: {e}")
            global_fallbacks = {}

        for col in all_columns:
            field_name = col.id
            header_key = col.header.strip() if getattr(col, "header", None) else ""
            
            # Skip structural columns
            if col.id in ["col_static", "col_qty_header", "col_no"]:
                self.logger.debug(f"    [Skipped]  {col.id} (Structural/Static)")
                continue
            
            fallback_val = None
            if self.fallback_mappings:
                if col.id in self.fallback_mappings:
                    fallback_val = self.fallback_mappings[col.id]
                elif header_key and header_key in self.fallback_mappings:
                    fallback_val = self.fallback_mappings[header_key]
                elif getattr(col, "header", None) and col.header in self.fallback_mappings:
                    fallback_val = self.fallback_mappings[col.header]

            if fallback_val is not None:
                mappings[field_name] = {"fallback": fallback_val}
                self.logger.info(f"    [Fallback] {col.id} fallback set to '{fallback_val}'")
                continue

            if col.id == "col_desc" and sheet.static_content_hints:
                dynamic_fallback = sheet.static_content_hints.get("description_fallback")
                if dynamic_fallback:
                    mappings[field_name] = {
                        "fallback": {
                            "standard": dynamic_fallback,
                            "daf": dynamic_fallback
                        }
                    }
                    self.logger.info(f"    [Override] col_desc fallback set to '{dynamic_fallback}'")
                    continue

            # Check global fallback strategies from DB if column not already in mappings/overridden
            if global_fallbacks:
                g_fallback = None
                if header_key and header_key in global_fallbacks:
                    g_fallback = global_fallbacks[header_key]
                elif getattr(col, "header", None) and col.header in global_fallbacks:
                    g_fallback = global_fallbacks[col.header]
                elif col.id in global_fallbacks:
                    g_fallback = global_fallbacks[col.id]

                if g_fallback is not None and g_fallback != "":
                    mappings[field_name] = {"fallback": g_fallback}
                    self.logger.info(f"    [Global Fallback] {col.id} fallback set to '{g_fallback}'")
                    continue
            
            # Skip columns already covered by defaults (no override to emit)
            if col.id in self._default_mapping_keys:
                self.logger.debug(f"    [Inherit]  {col.id} (covered by defaults, not emitted)")
                continue
            
            # Column NOT in defaults: emit empty {} so data_preparer processes it
            mappings[field_name] = {}
            self.logger.info(f"    [Auto]     {col.id} (identity mapping)")
        
        return {
            "_comment": "Inherits from defaults. Only overrides go here.",
            "mappings": mappings
        }

    
    def _build_static_payload(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build flat static_payload dictionary for a sheet."""
        static_payload = {}
        
        all_col_ids = [c.id for c in sheet.columns]
        for c in sheet.columns:
            all_col_ids.extend([child.id for child in c.children])

        if "col_static" in all_col_ids:
            static_lines = []
            if sheet.static_content_hints:
                static_lines = sheet.static_content_hints.get("static_lines", [])
            static_payload["col_static"] = static_lines

        if sheet.static_content_hints and sheet.static_content_hints.get("description_fallback"):
            static_payload["description_fallback"] = sheet.static_content_hints.get("description_fallback")
        
        return static_payload
    
    def _build_hs_code(self, sheet: SheetAnalysis) -> Optional[Dict[str, Any]]:
        """
        Build HS Code metadata block if present in template analysis.
        """
        is_contract = "contract" in sheet.name.lower()
        if sheet.footer_info and sheet.footer_info.has_hs_code and not is_contract:
            return {
                "col_id": sheet.footer_info.hs_code_col_id or "col_po",
                "value": sheet.footer_info.hs_code_text or "",
                "colspan": sheet.footer_info.hs_code_colspan,
                "style_context": "hs_code"
            }
        return None

    def _build_footer(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Delegates to the extracted footer builder module."""
        return build_footer(sheet)

    def _build_summary(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Delegates to the extracted summary builder module."""
        return build_summary(sheet)




if __name__ == "__main__":
    import sys
    import json
    from .scanner import WorkbookManager
    
    from core.logger_config import setup_logging
    from core.system_config import sys_config
    setup_logging(log_dir=sys_config.run_log_dir)
    
    if len(sys.argv) < 2:
        print("Usage: python config_builder.py <template.xlsx>")
        sys.exit(1)
    
    scanner = WorkbookManager()
    result = scanner.scan_template(sys.argv[1])
    
    builder = ConfigBuilder()
    bundle = builder.build_config(result)
    
    print(json.dumps(bundle, indent=2, ensure_ascii=False))
