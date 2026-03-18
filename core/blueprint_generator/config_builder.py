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

from .excel_scanner import TemplateAnalysisResult, SheetAnalysis, ColumnInfo
from .validator import BlueprintLogicValidator
from core.utils.snitch import snitch

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
            "processing": self._build_processing(analysis),
            "styling_bundle": self._build_styling_bundle(analysis),
            "layout_bundle": self._build_layout_bundle(analysis)
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
        
        # [Verification] Strict Mode: Ensure all IDs exist in Mapper Config
        # Delegated to BlueprintLogicValidator
        BlueprintLogicValidator.verify_strict_mode(sheet)
        
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

        for col in all_columns:
            field_name = col.id
            
            # Skip structural columns
            if col.id in ["col_static", "col_qty_header", "col_no"]:
                self.logger.debug(f"    [Skipped]  {col.id} (Structural/Static)")
                continue
            
            # [Smart Feature] Inject Dynamic Description Fallback (sheet-specific override)
            if col.id == "col_desc" and sheet.static_content_hints:
                dynamic_fallback = sheet.static_content_hints.get("description_fallback")
                if dynamic_fallback:
                    if dynamic_fallback == "LEATHER":
                        dynamic_fallback = "COW LEATHER"
                    
                    mappings[field_name] = {
                        "fallback": {
                            "standard": dynamic_fallback,
                            "daf": dynamic_fallback
                        }
                    }
                    self.logger.info(f"    [Override] col_desc fallback set to '{dynamic_fallback}'")
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
        """
        Build footer section for a sheet.
        
        Only emits sheet-specific footer data (merge_rules, add_ons).
        sum_cols and footer_cells are inherited from layout_bundle.defaults.footer.
        """
        # --- Detect merge rules from scanner ---
        merge_rules = []
        
        if sheet.footer_info:
            self.logger.info(f"  [Smart] Using detected footer info for {sheet.name}")
            total_col = sheet.footer_info.total_text_col_id
            
            # Create merge rule if colspan > 1
            if sheet.footer_info.merge_curr_colspan > 1:
                merge_rules.append({
                    "start_column_id": total_col,
                    "colspan": sheet.footer_info.merge_curr_colspan,
                    "comment": "Auto-detected from template"
                })
                self.logger.info(f"    [Smart] Added merge rule: {total_col} spans {sheet.footer_info.merge_curr_colspan} columns")
        else:
            self.logger.warning(f"  ⚠ No footer 'TOTAL' text found for {sheet.name}. Check template footer row.")
        
        footer = {
            "_comment": "Inherits sum_cols from defaults. footer_cells detected per-sheet.",
            "merge_rules": merge_rules,
            "add_ons": self._build_footer_addons(sheet)
        }
        
        # Build per-sheet footer_cells from detected FooterInfo
        if sheet.footer_info:
            footer_cells = []
            # Add TOTAL label cell (e.g. ["TOTAL OF:", "col_no"])
            if sheet.footer_info.total_text_col_id:
                footer_cells.append([
                    sheet.footer_info.total_text,
                    sheet.footer_info.total_text_col_id
                ])
                self.logger.info(f"    [Smart] footer_cells: TOTAL label '{sheet.footer_info.total_text}' -> {sheet.footer_info.total_text_col_id}")
            # Add pallet count cell only if detected in this sheet's template
            if sheet.footer_info.pallet_count_col_id:
                footer_cells.append([
                    "{pallet_count} PALLETS",
                    sheet.footer_info.pallet_count_col_id
                ])
                self.logger.info(f"    [Smart] footer_cells: pallet count -> {sheet.footer_info.pallet_count_col_id}")
            if footer_cells:
                footer["footer_cells"] = footer_cells
        
        return footer
    
    def _build_footer_addons(self, sheet: SheetAnalysis) -> Dict[str, Any]:
        """Build footer add-ons configuration."""
        add_ons = {}
        
        # before_footer (HS.CODE line)
        has_hs_code = False
        hs_code_text = "HS.CODE: 4107.12.00"
        hs_code_colspan = 1

        if sheet.footer_info:
            has_hs_code = sheet.footer_info.has_hs_code
            if sheet.footer_info.hs_code_text:
                hs_code_text = sheet.footer_info.hs_code_text
            hs_code_colspan = sheet.footer_info.hs_code_colspan
            
        add_ons["before_footer"] = {
            "enabled": True,
            "column_id": "col_po",
            "text": hs_code_text
        }
        
        if hs_code_colspan > 1:
            add_ons["before_footer"]["merge"] = hs_code_colspan
        
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



if __name__ == "__main__":
    import sys
    import json
    from .excel_scanner import ExcelLayoutScanner
    
    from core.logger_config import setup_logging
    from core.system_config import sys_config
    setup_logging(log_dir=sys_config.run_log_dir)
    
    if len(sys.argv) < 2:
        print("Usage: python config_builder.py <template.xlsx>")
        sys.exit(1)
    
    scanner = ExcelLayoutScanner()
    result = scanner.scan_template(sys.argv[1])
    
    builder = ConfigBuilder()
    bundle = builder.build_config(result)
    
    print(json.dumps(bundle, indent=2, ensure_ascii=False))
