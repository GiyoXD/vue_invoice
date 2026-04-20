from __future__ import annotations

import logging
import re
from typing import List, Optional, TYPE_CHECKING
from dataclasses import dataclass
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, Any

from .content_extractor import find_total_label_cell, find_pallet_count_column, find_footer_hs_code

if TYPE_CHECKING:
    from core.blueprint_generator.excel_scanner import ColumnInfo

logger = logging.getLogger(__name__)

@dataclass
class FooterInfo:
    """Information about the footer structure."""
    row_num: int
    total_text: str
    total_text_col_id: str
    merge_curr_colspan: int
    pallet_count_col_id: Optional[str] = None
    has_hs_code: bool = False
    hs_code_text: Optional[str] = None
    hs_code_colspan: int = 1
    hs_code_col_id: Optional[str] = None
    is_exact: bool = True


def find_column_id_by_index(col_index: int, columns: List[ColumnInfo]) -> Optional[str]:
    """
    Map a raw Excel column index (1-based) to its corresponding column ID.
    
    Checks both exact match and colspan range (for merged header columns).
    Returns None if no matching column is found.
    
    Args:
        col_index: 1-based Excel column index.
        columns: List of ColumnInfo from the scanner.
        
    Returns:
        The column ID string (e.g. 'col_desc', 'col_pallet_count'), or None.
    """
    for col in columns:
        if col.col_index == col_index:
            return col.id
        if col.col_index <= col_index < col.col_index + col.colspan:
            return col.id
    logger.warning(f"    ⚠ Column index {col_index} not covered by any detected column. Check template header detection.")
    return None


def scan_footer(worksheet: Worksheet, header_row: int, columns: List[ColumnInfo], logger: logging.Logger, sheet_name: str = "Unknown", mapping_config: Optional[Dict[str, Any]] = None) -> Optional[FooterInfo]:
    """
    Analyze the footer structure by searching for 'TOTAL' and 'X PALLETS'.
    
    Args:
        worksheet: The worksheet to scan.
        header_row: The header row number (scan starts below this).
        columns: List of ColumnInfo from the scanner.
        logger: Logger instance for output.
        sheet_name: Name of the sheet being scanned (for log context).
        mapping_config: Optional global mapping config containing 'footer_label_mappings'.
        
    Returns:
        FooterInfo if found, or None.
    """
    start_scan = header_row + 1
    end_scan = min(worksheet.max_row, header_row + 500)
    
    # --- Step 1: Find the TOTAL label cell ---
    result = find_total_label_cell(worksheet, start_scan, end_scan, mapping_config, logger, sheet_name)
    if not result:
        return None
        
    found_cell, is_exact = result
    
    # --- Step 2: Determine merge colspan at the TOTAL cell ---
    colspan = _get_cell_merge_colspan(worksheet, found_cell)
    
    # --- Step 3: Map TOTAL cell position to column ID ---
    total_col_id = find_column_id_by_index(found_cell.column, columns)
    if not total_col_id:
        logger.warning(f"    ⚠ [{sheet_name}] TOTAL label at column {found_cell.column} could not be mapped to a column ID.")
    
    # --- Step 4: Find pallet count column on the same row ---
    pallet_col_id = find_pallet_count_column(worksheet, found_cell.row, columns, find_column_id_by_index, logger, sheet_name)
    
    # --- Step 5: Find HS Code row ---
    hs_code_text, hs_code_colspan, hs_code_col_idx = find_footer_hs_code(worksheet, start_scan, end_scan)
    hs_code_col_id = None
    if hs_code_col_idx:
        hs_code_col_id = find_column_id_by_index(hs_code_col_idx, columns)

    return FooterInfo(
        row_num=found_cell.row,
        total_text=str(found_cell.value).strip(),
        total_text_col_id=total_col_id,
        merge_curr_colspan=colspan,
        pallet_count_col_id=pallet_col_id,
        has_hs_code=bool(hs_code_text),
        hs_code_text=hs_code_text,
        hs_code_colspan=hs_code_colspan,
        hs_code_col_id=hs_code_col_id,
        is_exact=is_exact
    )





def _get_cell_merge_colspan(worksheet: Worksheet, cell) -> int:
    """
    Check if a cell is part of a merged range and return the colspan.
    
    Returns:
        The number of columns spanned (1 if not merged).
    """
    for merged in worksheet.merged_cells.ranges:
        if merged.min_row <= cell.row <= merged.max_row:
            if merged.min_col <= cell.column <= merged.max_col:
                return merged.max_col - merged.min_col + 1
    return 1


