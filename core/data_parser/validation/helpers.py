import decimal
from typing import Any
from ..config import HEADER_SEARCH_COL_RANGE

# Set precision for Decimal calculations (consistent with data_processor)
decimal.getcontext().prec = 28


def _find_first_data_row(sheet, header_row: int, max_scan: int = 10) -> int:
    """Finds the first row index below header_row that contains actual table data."""
    for r in range(header_row + 1, min(header_row + 1 + max_scan, sheet.max_row + 1)):
        has_data = False
        for c in range(HEADER_SEARCH_COL_RANGE[0], HEADER_SEARCH_COL_RANGE[1] + 1):
            val = sheet.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                has_data = True
                break
        if has_data:
            return r
    return header_row + 1


def _is_valid_cell_value(val: Any, is_numeric: bool = False) -> bool:
    """Checks if a cell value is present and non-zero (if numeric)."""
    if val is None:
        return False
    val_str = str(val).strip()
    if not val_str:
        return False
    if not is_numeric:
        return True
    try:
        return decimal.Decimal(val_str.replace(',', '')) != 0
    except (decimal.InvalidOperation, ValueError, TypeError):
        try:
            from ..data_processor.cbm import _calculate_single_cbm
            cbm_val = _calculate_single_cbm(val_str, 0)
            return cbm_val is not None and cbm_val != 0
        except Exception:
            return False
