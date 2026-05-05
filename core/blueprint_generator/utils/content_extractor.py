import logging
import re
from typing import List, Optional, Tuple, Dict, Any, TYPE_CHECKING
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import MergedCell

if TYPE_CHECKING:
    from core.blueprint_generator.excel_scanner import ColumnInfo

logger = logging.getLogger(__name__)

# --- REGEX PATTERNS ---
# Description label pattern: matches "DES:", "DESC. :", "DESCRIPTION:", etc.
DESC_LABEL_PATTERN = re.compile(r'^(?:DES|DESC|DESCRIPTION|DES\.|DESC\.)\s*[:：]\s*(.+)$', re.IGNORECASE)

# Pallet count literal pattern: matches "25 PALLETS", "1 PALLET"
PALLET_PATTERN = re.compile(r'\d+\s*PALLETS?', re.IGNORECASE)
# Pallet count formula pattern: matches "=SUM(...) & \" PALLETS\""
PALLET_FORMULA_PATTERN = re.compile(r'PALLETS?', re.IGNORECASE)

from core.utils.loop_profiler import tick, loop_profiler
from core.blueprint_generator.rules import BlueprintRules

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
                capped_max_col = min(max_col, max(50, BlueprintRules.MAX_SCAN_COLUMN))
                
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
            for row in range(header_row + 1, min(header_row + max_sample_rows + 1, worksheet.max_row + 1)):
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


# --- 2. TABLE DESCRIPTION FALLBACK (secondary path) ---

def extract_table_fallback_description(worksheet: Worksheet, data_start: int, data_end: int, col_desc_index: int) -> Optional[str]:
    """
    Secondary fallback: finds description category labels from col_desc.

    Strategy: look for vertically-merged cells in col_desc — these are category
    labels (e.g. "Leather", "COW LEATHER") that span multiple data rows.
    Single-row values are specific product names and are ignored.

    Scans a bounded window (data_start → min(data_end, data_start+100)) and stops
    at the first TOTAL/SUB-TOTAL-like row to avoid picking up footer content.
    """
    TOTAL_KEYWORDS = {"TOTAL", "SUBTOTAL", "SUB TOTAL", "GRAND TOTAL", "AMOUNT"}
    MAX_SCAN = 100  # Cap to avoid reading bank info / legal text deep in the sheet

    # Build a lookup of merged ranges in col_desc column
    merged_spans: list[tuple[int, int, str]] = []  # (min_row, max_row, value)
    for merged in worksheet.merged_cells.ranges:
        if merged.min_col <= col_desc_index <= merged.max_col and merged.max_row > merged.min_row:
            if merged.min_row >= data_start:
                val = _get_cell_value_safe(worksheet, worksheet.cell(row=merged.min_row, column=col_desc_index))
                if val:
                    merged_spans.append((merged.min_row, merged.max_row, str(val).strip()))

    # Sort by span length descending (longest = most prominent category)
    merged_spans.sort(key=lambda x: x[1] - x[0], reverse=True)

    # Filter: stop at footer boundary, skip very long text (bank info / legal notes)
    scan_end = min(data_end, data_start + MAX_SCAN)
    categories = []
    seen = set()
    for min_row, max_row, val in merged_spans:
        if min_row > scan_end:
            continue
        upper = val.upper()
        if any(kw in upper for kw in TOTAL_KEYWORDS):
            continue
        if len(val) > 80:  # skip long legal/bank text
            continue
        if val not in seen:
            categories.append(val)
            seen.add(val)

    if categories:
        result = " / ".join(categories)
        logger.info(f"    [Extracted] Fallback description from merged col_desc cells: '{result}'")
        return result
    return None





@loop_profiler.watch("content_extractor.find_footer_hs_code")
def find_footer_hs_code(worksheet: Worksheet, start_row: int, end_row: int) -> Tuple[Optional[str], int, Optional[int]]:
    """
    Scan specifically in the footer bounds for HS Code to determine if it's there, its colspan, and its column.
    """
    hs_keywords = {"HS.CODE", "HS CODE", "HS-CODE", "H.S. CODE", "H.S CODE", "H.S.CODE", "HS. CODE"}
    
    for row in range(start_row, end_row + 1):
        for col in range(1, min(worksheet.max_column + 1, BlueprintRules.MAX_SCAN_COLUMN)):
            tick("content_extractor.find_footer_hs_code", sub="cells_scanned")
            cell = worksheet.cell(row=row, column=col)
            val = _get_cell_value_safe(worksheet, cell)
            if not val:
                continue
                
            upper_val = val.upper()
            if any(kw in upper_val for kw in hs_keywords):
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
        for col in range(1, min(worksheet.max_column + 1, BlueprintRules.MAX_SCAN_COLUMN)):
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
    for col in range(1, min(worksheet.max_column + 1, BlueprintRules.MAX_SCAN_COLUMN)):
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
