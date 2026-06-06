import logging
from typing import Dict, List, Any, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet

from core.blueprint_generator.schema import BlueprintSchema
from core.utils.loop_profiler import loop_profiler, tick
from core.blueprint_generator.utils.footer_scanner import FooterInfo, scan_footer
from core.blueprint_generator.utils.content_extractor import (
    detect_static_description_label,
    extract_static_column_values
)
from .models import SheetAnalysis, ColumnInfo
from .header_detector import HeaderDetector, get_cell_value
from .column_analyzer import ColumnAnalyzer

logger = logging.getLogger(__name__)


class SheetAnalyzer:
    """Orchestrates sheet-level analysis and properties extraction."""

    def __init__(self, header_detector: HeaderDetector, column_analyzer: ColumnAnalyzer):
        self.header_detector = header_detector
        self.column_analyzer = column_analyzer
        self.logger = logger

    @loop_profiler.watch("scanner._analyze_sheet")
    def analyze_sheet(self, worksheet: Worksheet, sheet_name: str, 
                       mapping_config: Optional[Dict[str, Any]] = None,
                       skip_desc_scan: bool = False,
                       skip_hs_scan: bool = False) -> Optional[SheetAnalysis]:
        """Analyze a single worksheet."""
        try:
            self.logger.info(f"  Analyzing sheet: {sheet_name}")
            # Filter out unsupported sheets before scanning
            normalized_name = sheet_name.lower().strip()
            
            # Fast mapping resolution using nested structure
            if mapping_config and isinstance(mapping_config, dict):
                sheet_mappings = mapping_config.get('sheet_name_mappings', {}).get('mappings', {})
                if isinstance(sheet_mappings, dict):
                    # Fast case-insensitive exact matching
                    lower_mappings = {k.lower().strip(): v for k, v in sheet_mappings.items()}
                    if normalized_name in lower_mappings:
                        normalized_name = lower_mappings[normalized_name].lower().strip()

            is_supported = False
            
            # Create a set of variants for matching (with/without underscores/spaces)
            variants_to_check = {
                normalized_name,
                normalized_name.replace(' ', '_'),
                normalized_name.replace('_', ' ')
            }
            
            # 1. Exact match check (important for mapped system names like "summary_packing_list")
            for variant in variants_to_check:
                if variant in BlueprintSchema.ALLOWED_SEARCH_SHEETS:
                    is_supported = True
                    break
                    
            if not is_supported:
                self.logger.info(f"    Skipping sheet '{sheet_name}': Not in allowed search list.")
                return None
            
            # Find header row
            header_row, header_cells = self.header_detector.find_header_row(worksheet, mapping_config)
            if not header_row:
                self.logger.warning(f"    No header row found in {sheet_name}")
                return None
            
            self.logger.info(f"    Header row: {header_row}")
            
            # [Smart Feature] Check for multi-row headers and determine data start row early
            # This is critical for accurate data sampling and style extraction.
            has_multi_row = self.check_multi_row_header(worksheet, header_row)
            data_start_row = header_row + 1
            if has_multi_row:
                 # Find the bottom-most row of the header (max span of merges starting at header_row)
                 for merged in worksheet.merged_cells.ranges:
                     if merged.min_row == header_row:
                          data_start_row = max(data_start_row, merged.max_row + 1)
            
            self.logger.info(f"    Header type: {'Multi-row' if has_multi_row else 'Single-row'}. Data starts at row: {data_start_row}")

            # Analyze columns
            columns = self.column_analyzer.analyze_columns(worksheet, header_row, header_cells, data_start_row, mapping_config)
            self.logger.info(f"    Found {len(columns)} columns")
            
            # Determine data source type
            data_source = self.determine_data_source(sheet_name, columns, mapping_config)
            self.logger.info(f"    Data source: {data_source}")
            
            # Extract font info
            header_font = self.extract_font_info(worksheet, header_row, 1)
            data_font = self.extract_font_info(worksheet, data_start_row, 1)
            
            # Extract row heights
            row_heights = self.extract_row_heights(worksheet, header_row, data_source, data_start_row)
            
            # [Smart Feature] Dynamic Footer Analysis (Delegated to Utility)
            footer_info = scan_footer(worksheet, header_row, columns, self.logger, sheet_name=sheet_name, mapping_config=mapping_config, skip_hs_scan=skip_hs_scan)
            if footer_info:
                self.logger.info(f"    Footer detected at row {footer_info.row_num}: '{footer_info.total_text}' (colspan={footer_info.merge_curr_colspan})")
            
            # Detect static content hints (like "Mark & Nº" column content) strictly bounded by the total row anchor
            footer_row = footer_info.row_num if footer_info else None
            static_hints = self.detect_static_content(worksheet, header_row, columns, skip_desc_scan, footer_row=footer_row)
            
            sheet_analysis = SheetAnalysis(
                name=sheet_name,
                header_row=header_row,
                columns=columns,
                data_source=data_source,
                header_font=header_font,
                data_font=data_font,
                row_heights=row_heights,
                has_multi_row_header=has_multi_row,
                static_content_hints=static_hints,
                footer_info=footer_info
            )
            
            # [Proactive Warning] Check for missing footer on financial/aggregation sheets
            if not footer_info and data_source == "aggregation":
                warning_msg = (
                    f"[{sheet_name}] ⚠️ Footer (Total row) NOT detected. "
                    f"WHAT TO DO: Ensure the sheet has a row starting with 'TOTAL' or 'TOTAL AMOUNT'. "
                    f"If the label is different, update 'footer_scanner.py' or 'mapping_config.json'."
                )
                self.logger.warning(warning_msg)
                setattr(sheet_analysis, "_temp_warning", warning_msg)

            return sheet_analysis
            
        except Exception as e:
            self.logger.error(f"    Error analyzing {sheet_name}: {e}")
            return None

    def determine_data_source(self, sheet_name: str, columns: List[ColumnInfo], mapping_config: Optional[Dict[str, Any]] = None) -> str:
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
    
    def extract_font_info(self, worksheet: Worksheet, row: int, col: int) -> Dict[str, Any]:
        """
        Extract font information from a cell.
        """
        cell = worksheet.cell(row=row, column=col)
        font = cell.font

        if not font:
            raise ValueError(f"No font detected at Row {row}, Col {col}. Please ensure the template cell has explicit styling.")

        name = font.name
        size = font.size

        if name is None: name = "Calibri"
        if size is None: size = 11.0

        return {
            "name": name,
            "size": size,
            "bold": font.bold or False,
            "italic": font.italic or False
        }
    
    def extract_row_heights(self, worksheet: Worksheet, header_row: int, data_source: str = "dataset_default", data_start_row: int = None) -> Dict[str, float]:
        """Extract row heights using 3-step fallback strategy."""
        if data_start_row is None:
            data_start_row = header_row + 1
        
        def get_height(row_idx: int, standard_key: str) -> float:
            # 1. Explicit
            if row_idx in worksheet.row_dimensions:
                h = worksheet.row_dimensions[row_idx].height
                if h is not None:
                    return float(h)
            
            # 2. Sheet Default (Strict: If 15.0 or None, fail)
            default_h = worksheet.sheet_format.defaultRowHeight
            if default_h is not None and default_h != 15.0:
                 return float(default_h)
                 
            # 3. Failure
            self.logger.warning(f"Row {row_idx} has no explicit height. Defaulting to 20.0")
            return 20.0

        try:
            header_height = get_height(header_row, "header")
        except ValueError as e:
             self.logger.warning(f"Header Row {header_row} height detection failed: {e}")
             header_height = 20.0
            
        # Scan next 10 rows for data height (Median)
        heights = []
        for r in range(data_start_row, min(data_start_row + 10, worksheet.max_row + 1)):
            heights.append(get_height(r, "data"))
        
        if heights:
            heights.sort()
            mid = len(heights) // 2
            data_height = heights[mid]
        else:
            data_height = get_height(data_start_row, "data")
            
        return {
            "header": header_height,
            "data": data_height,
            "footer": header_height  # Usually same as header
        }
    
    @loop_profiler.watch("scanner._check_multi_row_header")
    def check_multi_row_header(self, worksheet: Worksheet, header_row: int) -> bool:
        """Check if there's a multi-row header structure."""
        for merged in worksheet.merged_cells.ranges:
            tick("scanner._check_multi_row_header", sub="merge_ranges_checked")
            if merged.min_row == header_row and merged.max_row > header_row:
                return True
        return False
    
    @loop_profiler.watch("scanner._detect_static_content")
    def detect_static_content(self, worksheet: Worksheet, header_row: int, 
                               columns: List[ColumnInfo], skip_desc_scan: bool = False,
                               footer_row: Optional[int] = None) -> Dict[str, List[str]]:
        """Detect static content patterns in the data area."""
        hints = {}
        
        if not skip_desc_scan:
            desc_fallback = detect_static_description_label(worksheet, header_row, columns)
            if desc_fallback:
                hints["description_fallback"] = desc_fallback
        
        static_lines = extract_static_column_values(worksheet, header_row, columns, footer_row=footer_row)
        if static_lines:
            hints["static_lines"] = static_lines
        
        return hints
