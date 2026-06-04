import logging
import re
from typing import List, Optional, Tuple, Dict, Any, TYPE_CHECKING
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import MergedCell

if TYPE_CHECKING:
    from core.blueprint_generator.internal.scanner import ColumnInfo

logger = logging.getLogger(__name__)

# --- REGEX PATTERNS ---
# Description label pattern: matches "DES:", "DESC. :", "DESCRIPTION:", etc.
DESC_LABEL_PATTERN = re.compile(r'^(?:DES|DESC|DESCRIPTION|DES\.|DESC\.)\s*[:：]\s*(.+)$', re.IGNORECASE)

# Pallet count literal pattern: matches "25 PALLETS", "1 PALLET"
PALLET_PATTERN = re.compile(r'\d+\s*PALLETS?', re.IGNORECASE)
# Pallet count formula pattern: matches "=SUM(...) & \" PALLETS\""
PALLET_FORMULA_PATTERN = re.compile(r'PALLETS?', re.IGNORECASE)

from core.utils.loop_profiler import tick, loop_profiler
from core.blueprint_generator.schema import BlueprintSchema

# --- HELPER ---
@loop_profiler.watch("content_extractor.get_cell_merge_colspan")
def get_cell_merge_colspan(worksheet: Worksheet, cell) -> int:
    """
    Check if a cell is part of a merged range and return the colspan.
    
    Returns:
        The number of columns spanned (1 if not merged).
    """
    for merged in worksheet.merged_cells.ranges:
        tick("content_extractor.get_cell_merge_colspan", sub="merge_ranges_scanned")
        if merged.min_row <= cell.row <= merged.max_row:
            if merged.min_col <= cell.column <= merged.max_col:
                return merged.max_col - merged.min_col + 1
    return 1

@loop_profiler.watch("content_extractor._get_cell_value_safe")
def _get_cell_value_safe(worksheet: Worksheet, cell) -> Optional[str]:
    """Safely get string value from a cell, handling MergedCells using an O(1) cache."""
    if cell.value is not None and not isinstance(cell, MergedCell):
        return str(cell.value)

    # Lazy-init a sparse column-indexed cache to avoid O(N*M) memory explosion
    if not hasattr(worksheet, '_merged_cells_cache'):
        cache = {} # col -> list of (min_row, max_row, val)
        for merged_range in worksheet.merged_cells.ranges:
            tick("content_extractor._get_cell_value_safe", sub="cache_build_iterations")
            min_col, min_row, max_col, max_row = merged_range.bounds
            # The top-left cell holds the actual value for the entire merged range
            top_left_cell = worksheet.cell(row=min_row, column=min_col)
            if top_left_cell.value is not None:
                val = str(top_left_cell.value)
                
                # Clamp column boundary to prevent huge lateral scans
                capped_max_col = min(max_col, max(50, BlueprintSchema.MAX_SCAN_COLUMN))
                
                # Expand only the columns (max ~50 iterations per merge, instead of 100,000)
                for c in range(min_col, capped_max_col + 1):
                    if c not in cache:
                        cache[c] = []
                    cache[c].append((min_row, max_row, val))
        worksheet._merged_cells_cache = cache
    else:
        tick("content_extractor._get_cell_value_safe", sub="cache_hits")

    # Sparse lookup: O(K) where K is number of merges in this specific column
    col_ranges = worksheet._merged_cells_cache.get(cell.column, [])
    for min_r, max_r, val in col_ranges:
        if min_r <= cell.row <= max_r:
            return val
            
    return None


# --- 1. FALLBACK DESCRIPTION EXTRACTION ---

@loop_profiler.watch("content_extractor.detect_static_description_label")
def detect_static_description_label(worksheet: Worksheet, header_row: int, columns: List['ColumnInfo'], max_sample_rows: int = 10) -> Optional[str]:
    """
    Detects the description label (e.g. "DES: COW LEATHER") from the static column.
    Uses regex for robust discovery.
    """
    for col in columns:
        tick("content_extractor.detect_static_description_label", sub="columns_checked")
        if col.id == "col_static":
            # Skip vertically merged header cells by finding their max row
            start_scan_row = header_row
            for merged in worksheet.merged_cells.ranges:
                if merged.min_row <= header_row <= merged.max_row:
                    if merged.min_col <= col.col_index <= merged.max_col:
                        start_scan_row = merged.max_row
                        break
                        
            for row in range(start_scan_row + 1, min(start_scan_row + max_sample_rows + 1, worksheet.max_row + 1)):
                tick("content_extractor.detect_static_description_label", sub="rows_sampled")
                cell = worksheet.cell(row=row, column=col.col_index)
                value = _get_cell_value_safe(worksheet, cell)
                if not value:
                    continue
                
                # Check using regex instead of strict startswith
                match = DESC_LABEL_PATTERN.search(value.strip())
                if match:
                    desc_part = match.group(1).strip()
                    if desc_part:
                        logger.info(f"    [Detection] Found description label in col_static using regex: '{desc_part}'")
                        return desc_part
    return None


@loop_profiler.watch("content_extractor.extract_static_column_values")
def extract_static_column_values(worksheet: Worksheet, header_row: int, columns: List['ColumnInfo'], footer_row: Optional[int] = None) -> List[str]:
    """
    Extracts the actual cell values from the static column ('col_static') starting below the header row.
    Returns the raw string values as-is.
    """
    col_static = next((col for col in columns if col.id == "col_static"), None)
    if not col_static:
        return []
        
    static_lines = []
    consecutive_empty = 0
    
    # Skip vertically merged header cells by finding their max row
    start_scan_row = header_row
    for merged in worksheet.merged_cells.ranges:
        if merged.min_row <= header_row <= merged.max_row:
            if merged.min_col <= col_static.col_index <= merged.max_col:
                start_scan_row = merged.max_row
                break
                
    # We scan down the static column starting below the header bounds, stopping strictly before the footer/total row
    limit_row = footer_row if footer_row is not None else (worksheet.max_row + 1)
    for row in range(start_scan_row + 1, min(start_scan_row + 16, limit_row)):
        cell = worksheet.cell(row=row, column=col_static.col_index)
        val = _get_cell_value_safe(worksheet, cell)
        
        if val is not None and str(val).strip():
            static_lines.append(str(val).strip())
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            
        # Stop scanning if we hit 3 consecutive empty rows
        if consecutive_empty >= 3:
            break
            
    return static_lines


# --- 2. TABLE DESCRIPTION FALLBACK (secondary path removed) ---





HS_CODE_PATTERN = re.compile(r'\bH\.?\s*S\.?\s*[-_.]?\s*C\s*O\s*D\s*E\b', re.IGNORECASE)


@loop_profiler.watch("content_extractor.find_footer_hs_code")
def find_footer_hs_code(worksheet: Worksheet, start_row: int, end_row: int) -> Tuple[Optional[str], int, Optional[int]]:
    """
    Scan specifically in the footer bounds for HS Code to determine if it's there, its colspan, and its column.
    """
    for row in range(start_row, end_row + 1):
        for col in range(1, min(worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
            tick("content_extractor.find_footer_hs_code", sub="cells_scanned")
            cell = worksheet.cell(row=row, column=col)
            val = _get_cell_value_safe(worksheet, cell)
            if not val:
                continue
                
            if HS_CODE_PATTERN.search(val):
                # Calculate colspan
                colspan = get_cell_merge_colspan(worksheet, cell)
                return val, colspan, col
                
    return None, 1, None



# --- 3. FOOTER ELEMENTS (PALLET & TOTAL LABELS) ---

@loop_profiler.watch("content_extractor.find_total_label_cell")
def find_total_label_cell(worksheet: Worksheet, start_row: int, end_row: int, mapping_config: Optional[Dict[str, Any]] = None, logger_instance: Optional[logging.Logger] = None, sheet_name: str = "Unknown"):
    """
    Scan rows for the first cell containing a TOTAL-like label.
    """
    total_keywords = []
    if mapping_config and "footer_label_mappings" in mapping_config:
        mappings = mapping_config["footer_label_mappings"].get("keywords", [])
        total_keywords = [kw.upper() for kw in mappings]
        
    if not total_keywords:
        if logger_instance:
            logger_instance.warning(f"    ⚠ [{sheet_name}] No footer keywords configured. Cannot detect footer row!")
        return None
    
    # Sort keywords by length descending to match most specific first (prevents "TOTAL" shadowing "GRAND TOTAL")
    total_keywords.sort(key=len, reverse=True)
    
    best_match = None
    
    for row in range(start_row, end_row + 1):
        for col in range(1, min(worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
            tick("content_extractor.find_total_label_cell", sub="cells_scanned")
            cell = worksheet.cell(row=row, column=col)
            val = _get_cell_value_safe(worksheet, cell)
            if not val:
                continue
                
            val_upper = val.upper().strip()
            for kw in total_keywords:
                if kw == val_upper:
                    # Found exact match: immediate return, confirmed.
                    return cell, True
                if kw in val_upper and not best_match:
                    # Found partial match: save as candidate if no exact match found later.
                    best_match = (cell, False)
    
    return best_match if best_match else None

@loop_profiler.watch("content_extractor.find_pallet_count_column")
def find_pallet_count_column(worksheet: Worksheet, footer_row: int, columns: List['ColumnInfo'], find_col_id_func, logger_instance: logging.Logger, sheet_name: str = "Unknown") -> Optional[str]:
    """
    Scan the footer row for a pallet count pattern.
    """
    for col in range(1, min(worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
        tick("content_extractor.find_pallet_count_column", sub="cols_scanned")
        cell = worksheet.cell(row=footer_row, column=col)
        val = _get_cell_value_safe(worksheet, cell)
        
        if not val:
            continue
        
        if PALLET_PATTERN.search(val) or (val.startswith("=") and PALLET_FORMULA_PATTERN.search(val)):
            pallet_col_id = find_col_id_func(col, columns)
            logger_instance.info(f"    [{sheet_name}] Pallet count detected at col {col} -> {pallet_col_id}")
            return pallet_col_id
    
    logger_instance.warning(f"    ⚠ [{sheet_name}] No pallet count pattern found on footer row {footer_row}")
    return None
