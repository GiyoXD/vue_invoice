from __future__ import annotations

import logging
import re
from typing import List, Optional, TYPE_CHECKING
from dataclasses import dataclass
from openpyxl.worksheet.worksheet import Worksheet

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


def scan_footer(worksheet: Worksheet, header_row: int, columns: List[ColumnInfo], logger: logging.Logger, sheet_name: str = "Unknown") -> Optional[FooterInfo]:
    """
    Analyze the footer structure by searching for 'TOTAL' and 'X PALLETS'.
    
    Args:
        worksheet: The worksheet to scan.
        header_row: The header row number (scan starts below this).
        columns: List of ColumnInfo from the scanner.
        logger: Logger instance for output.
        sheet_name: Name of the sheet being scanned (for log context).
        
    Returns:
        FooterInfo if found, or None.
    """
    start_scan = header_row + 1
    end_scan = min(worksheet.max_row, header_row + 500)
    
    # --- Step 1: Find the TOTAL label cell ---
    found_cell = _find_total_label_cell(worksheet, start_scan, end_scan)
    if not found_cell:
        return None
    
    # --- Step 2: Determine merge colspan at the TOTAL cell ---
    colspan = _get_cell_merge_colspan(worksheet, found_cell)
    
    # --- Step 3: Map TOTAL cell position to column ID ---
    total_col_id = find_column_id_by_index(found_cell.column, columns)
    if not total_col_id:
        logger.warning(f"    ⚠ [{sheet_name}] TOTAL label at column {found_cell.column} could not be mapped to a column ID.")
    
    # --- Step 4: Find pallet count column on the same row ---
    pallet_col_id = _find_pallet_count_column(worksheet, found_cell.row, columns, logger, sheet_name)
    
    # --- Step 5: Find HS Code row ---
    hs_code_text, hs_code_colspan = _find_hs_code(worksheet, start_scan, end_scan)

    return FooterInfo(
        row_num=found_cell.row,
        total_text=str(found_cell.value).strip(),
        total_text_col_id=total_col_id,
        merge_curr_colspan=colspan,
        pallet_count_col_id=pallet_col_id,
        has_hs_code=bool(hs_code_text),
        hs_code_text=hs_code_text,
        hs_code_colspan=hs_code_colspan
    )


def _find_total_label_cell(worksheet: Worksheet, start_row: int, end_row: int):
    """
    Scan rows for the first cell containing a TOTAL-like label.
    
    Matches: 'TOTAL', 'TOTAL:', 'TOTAL OF:', 'TOTAL OF', 'TOTAL：'
    
    Returns:
        The cell object if found, or None.
    """
    total_keywords = {"TOTAL", "TOTAL:", "TOTAL OF:", "TOTAL OF", "TOTAL："}
    
    for row in range(start_row, end_row + 1):
        for col in range(1, min(worksheet.max_column + 1, 20)):
            cell = worksheet.cell(row=row, column=col)
            val = str(cell.value).strip().upper() if cell.value else ""
            
            if val in total_keywords or val.startswith("TOTAL OF"):
                return cell
    return None


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


# Pattern for literal text like "25 PALLETS" or "1 PALLET"
PALLET_PATTERN = re.compile(r'\d+\s*PALLETS?', re.IGNORECASE)
# Pattern for formula-based pallet cells like '=SUM(D21:D25) & " PALLETS"'
PALLET_FORMULA_PATTERN = re.compile(r'PALLETS?', re.IGNORECASE)

def _find_pallet_count_column(worksheet: Worksheet, footer_row: int, columns: List[ColumnInfo], logger: logging.Logger, sheet_name: str = "Unknown") -> Optional[str]:
    """
    Scan the footer row for a pallet count pattern.
    
    Detects both:
    - Literal text: '25 PALLETS', '1 PALLET'
    - Formula-based: '=SUM(D21:D25) & " PALLETS"'
    
    Args:
        worksheet: The worksheet to scan.
        footer_row: The row number to scan.
        columns: List of ColumnInfo from the scanner.
        logger: Logger instance for output.
        sheet_name: Name of the sheet being scanned (for log context).
        
    Returns:
        The column ID where the pallet count was found, or None.
    """
    for col in range(1, min(worksheet.max_column + 1, 20)):
        cell = worksheet.cell(row=footer_row, column=col)
        val = str(cell.value).strip() if cell.value else ""
        
        if not val:
            continue
        
        # Check literal text first (e.g. "25 PALLETS")
        # Then check formula-based (e.g. '=SUM(...) & " PALLETS"')
        if PALLET_PATTERN.search(val) or (val.startswith("=") and PALLET_FORMULA_PATTERN.search(val)):
            pallet_col_id = find_column_id_by_index(col, columns)
            logger.info(f"    [{sheet_name}] Pallet count detected at col {col} -> {pallet_col_id}")
            return pallet_col_id
    
    logger.warning(f"    ⚠ [{sheet_name}] No pallet count pattern (e.g. '25 PALLETS') found on footer row {footer_row}")
    return None


def _find_hs_code(worksheet: Worksheet, start_row: int, end_row: int) -> tuple[Optional[str], int]:
    """
    Scan rows for a cell containing HS Code information.
    
    Matches variations of 'HS.CODE', 'HS CODE', 'HS-CODE'.
    
    Returns:
        A tuple of (text value if found else None, colspan of the merged cell if merged else 1).
    """
    hs_keywords = {"HS.CODE", "HS CODE", "HS-CODE"}
    
    for row in range(start_row, end_row + 1):
        for col in range(1, min(worksheet.max_column + 1, 20)):
            cell = worksheet.cell(row=row, column=col)
            val = str(cell.value).strip() if cell.value else ""
            upper_val = val.upper()
            
            # Check if any of the keywords are in the string
            if any(kw in upper_val for kw in hs_keywords):
                colspan = _get_cell_merge_colspan(worksheet, cell)
                return val, colspan
                
    return None, 1

