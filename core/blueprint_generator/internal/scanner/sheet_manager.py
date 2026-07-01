import logging
from typing import Dict, List, Any, Optional
from openpyxl.worksheet.worksheet import Worksheet

from core.blueprint_generator.schema import BlueprintSchema
from core.utils.loop_profiler import loop_profiler
from .models import SheetAnalysis, ColumnInfo
from .models.addons import LeatherSummaryFact
from .header_detector import BoundaryDetector
from .tabular_scanner import TabularScanner
from .template_scanner import TemplateScanner

logger = logging.getLogger(__name__)


class SheetManager:
    """Coordinates analysis of a single worksheet by sequencing zone scanners."""

    def __init__(self, boundary_detector: BoundaryDetector):
        self.boundary_detector = boundary_detector
        self.tabular_scanner = TabularScanner()
        self.template_scanner = TemplateScanner()
        self.logger = logger

    @loop_profiler.watch("scanner._analyze_sheet")
    def analyze_sheet(self, worksheet: Worksheet, sheet_name: str, 
                       mapping_config: Optional[Dict[str, Any]] = None,
                       skip_desc_scan: bool = False,
                       skip_hs_scan: bool = False) -> Optional[SheetAnalysis]:
        """Analyze a single worksheet using the 3-zone architecture."""
        try:
            self.logger.info(f"  Analyzing sheet: {sheet_name}")
            
            # Extract header text mappings from config
            header_mappings = None
            if mapping_config and isinstance(mapping_config, dict):
                header_mappings = mapping_config.get('header_text_mappings', {}).get('mappings', {})

            # 1. Detect boundaries
            boundaries = self.boundary_detector.detect_boundaries(
                worksheet, header_mappings, mapping_config, sheet_name
            )
            if not boundaries:
                self.logger.warning(f"    No header row found in {sheet_name}")
                return None
            
            self.logger.info(f"    Header row: {boundaries.header_row}")
            self.logger.info(f"    Header type: {'Multi-row' if boundaries.data_start_row > boundaries.header_row + 1 else 'Single-row'}. Data starts at row: {boundaries.data_start_row}")

            # 2. Scan table zone (Zone 2)
            table_layout = self.tabular_scanner.scan_table(
                worksheet, boundaries, mapping_config, skip_desc_scan, skip_hs_scan, sheet_name
            )
            self.logger.info(f"    Found {len(table_layout.columns)} columns")
            
            # 3. Classify data source type
            data_source = self._determine_data_source(sheet_name, mapping_config)
            self.logger.info(f"    Data source: {data_source}")
            
            # 4. Scan static zones (Zone 1 & 3)
            static_layout = self.template_scanner.scan_static_content(
                worksheet=worksheet, 
                boundaries=boundaries, 
                columns=table_layout.columns, 
                sheet_name=sheet_name
            )
            
            # The TemplateScanner only scans for layout right now, so static hints should be just what the table scanner found
            merged_hints = table_layout.static_content_hints
            
            # Detect leather summary rows directly from the openpyxl worksheet
            addon_facts = []
            if boundaries.footer_row and data_source == "processed_tables_multi":
                end_scan_row = boundaries.footer_end_row if boundaries.footer_end_row else boundaries.footer_row
                end_scan_row = max(end_scan_row, boundaries.footer_row + 5)
                end_scan_row = min(end_scan_row, worksheet.max_row)
                
                # Pre-compute column mappings
                def find_col_id(col_idx):
                    for col in table_layout.columns:
                        if col.col_index == col_idx:
                            return col.id
                        for child in col.children:
                            if child.col_index == col_idx:
                                return child.id
                    return None
                    
                for r in range(boundaries.footer_row + 1, end_scan_row + 1):
                    for c in range(1, min(worksheet.max_column + 1, 40)):
                        cell_val = str(worksheet.cell(row=r, column=c).value or "").strip()
                        if "total of:" in cell_val.lower():
                            next_c = c + 1
                            next_val = ""
                            while next_c <= min(worksheet.max_column, 40):
                                temp_val = str(worksheet.cell(row=r, column=next_c).value or "").strip()
                                if temp_val:
                                    next_val = temp_val
                                    break
                                next_c += 1
                                
                            if "leather" in next_val.lower():
                                next_val_lower = next_val.lower()
                                leather_key = "BUFFALO" if "buffalo" in next_val_lower else "COW"
                                
                                total_col_id = find_col_id(c) or "col_po"
                                label_col_id = find_col_id(next_c) or "col_item"
                                
                                fact = LeatherSummaryFact(
                                    leather_key=leather_key,
                                    total_col_id=total_col_id,
                                    label_col_id=label_col_id,
                                    total_value=cell_val,
                                    label_value=next_val
                                )
                                addon_facts.append(fact)
            merged_hints["addon_facts"] = addon_facts
            
            sheet_analysis = SheetAnalysis(
                name=sheet_name,
                header_row=boundaries.header_row,
                columns=table_layout.columns,
                data_source=data_source,
                header_font=table_layout.header_font,
                data_font=table_layout.data_font,
                row_heights=table_layout.row_heights,
                has_multi_row_header=table_layout.has_multi_row_header,
                static_content_hints=merged_hints,
                static_layout=static_layout,
                footer_info=table_layout.footer_info
            )
            
            # [Proactive Warning] Check for missing footer on financial/aggregation sheets
            if not table_layout.footer_info and data_source == "aggregation":
                warning_msg = (
                    f"[{sheet_name}] ⚠️ Footer (Total row) NOT detected. "
                    f"WHAT TO DO: Ensure the sheet has a row starting with 'TOTAL' or 'TOTAL AMOUNT'. "
                    f"If the label is different, update 'tabular_scanner.py' or 'mapping_config.json'."
                )
                self.logger.warning(warning_msg)
                setattr(sheet_analysis, "_temp_warning", warning_msg)

            return sheet_analysis
            
        except Exception as e:
            self.logger.error(f"    Error analyzing {sheet_name}: {e}")
            return None

    def _determine_data_source(self, sheet_name: str, 
                               mapping_config: Optional[Dict[str, Any]] = None) -> str:
        """
        Determine if this sheet is 'aggregation' (single table) 
        or 'processed_tables_multi' (repeating tables).
        """
        normalized_name = sheet_name.lower().strip()
        
        if mapping_config:
            sheet_mappings = mapping_config.get('sheet_name_mappings', {}).get('mappings', {})
            # Case-insensitive resolution
            lower_mappings = {k.lower().strip(): v for k, v in sheet_mappings.items()}
            if normalized_name in lower_mappings:
                normalized_name = lower_mappings[normalized_name].lower()
                
        variants_to_check = {
            normalized_name,
            normalized_name.replace(' ', '_'),
            normalized_name.replace('_', ' ')
        }
        
        # 1. Exact match check against BlueprintSchema definitions
        for variant in variants_to_check:
            if variant in BlueprintSchema.AGGREGATION_SHEETS:
                if variant.replace(' ', '_') == "summary_packing_list":
                    return "summary_packing_list"
                return "aggregation"
            elif variant in BlueprintSchema.PROCESSED_TABLES_SHEETS:
                return "processed_tables_multi"
            
        return "aggregation" # default   
