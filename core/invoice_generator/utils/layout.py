from openpyxl.worksheet.worksheet import Worksheet
from typing import List, Dict, Any, Optional, Tuple
from openpyxl.utils import get_column_letter
from ..styling.models import StylingConfigModel
from ..styling.style_applier import apply_cell_style, apply_header_style
from ..styling.style_config import THIN_BORDER, NO_BORDER, CENTER_ALIGNMENT, LEFT_ALIGNMENT, BOLD_FONT
from decimal import Decimal, InvalidOperation
import re
import traceback
import logging

logger = logging.getLogger(__name__)




def apply_column_widths(worksheet: Worksheet, sheet_styling_config: Optional[StylingConfigModel], header_map: Optional[Dict[str, int]]):
    """
    Sets column widths based on the configuration.

    Args:
        worksheet: The openpyxl Worksheet object.
        sheet_styling_config: Styling configuration containing the 'column_widths' dictionary.
        header_map: Dictionary mapping header text to column index (1-based).
    """
    if not sheet_styling_config or not header_map: return
    column_widths_cfg = sheet_styling_config.columnIdWidths
    if not column_widths_cfg or not isinstance(column_widths_cfg, dict): return
    for header_text, width in column_widths_cfg.items():
        col_idx = header_map.get(header_text)
        if col_idx:
            col_letter = get_column_letter(col_idx)
            try:
                width_val = float(width)
                if width_val > 0: worksheet.column_dimensions[col_letter].width = width_val
                else: pass # Ignore non-positive widths
            except (ValueError, TypeError): pass # Ignore invalid width values
            except Exception as width_err: pass # Log other errors?
        else: pass # Header text not found in map


def _estimate_display_text(cell_value, number_format: str) -> str:
    """
    Estimates the display text of a cell value as Excel would render it.

    For numbers, formats with comma grouping and 4 decimal places.
    For text values, returns str(cell_value) as-is.

    Args:
        cell_value: The raw cell value.
        number_format: The cell's number_format string (unused, kept for API).

    Returns:
        The estimated display text string.
    """
    if cell_value is None:
        return ""

    if isinstance(cell_value, str):
        return cell_value

    if isinstance(cell_value, (int, float)):
        try:
            return f"{cell_value:,.2f}"
        except (ValueError, TypeError):
            pass

    return str(cell_value)


def auto_fit_dimensions(
    worksheet: Worksheet,
    header_start_row: int,
    data_end_row: int,
    num_columns: int,
    padding: int = 5,
    line_height: float = 15.0,
    min_width: float = 8.0,
    max_width: float = 60.0
):
    """
    Auto-fits column widths and row heights based on actual cell content.

    Width strategy:
        For each column, scan data rows to find the longest formatted
        display value. Set width = len(formatted_value) * 1.3 + padding.

    Height strategy:
        For each row, count line breaks (\\n) in every cell.
        Set row height = max(line_breaks_in_row) * line_height.
        Only overrides rows where a cell has more than 1 line.

    Args:
        worksheet: The openpyxl Worksheet.
        header_start_row: First data row to scan.
        data_end_row: Last row to scan (last data or footer row).
        num_columns: Total number of columns to scan.
        padding: Extra character units added to the computed width.
        line_height: Points per line for height calculation.
        min_width: Minimum column width to prevent tiny columns.
        max_width: Maximum column width to prevent absurdly wide columns.
    """
    if header_start_row <= 0 or data_end_row <= 0 or num_columns <= 0:
        logger.warning(f"auto_fit_dimensions: invalid bounds (header_start={header_start_row}, data_end={data_end_row}, cols={num_columns})")
        return

    logger.info(f"auto_fit_dimensions: scanning rows {header_start_row}-{data_end_row}, {num_columns} columns")

    # Track the widest display text length per column + whether it's text or number
    # Format: {col_idx: (max_len, is_text)}
    col_max_info: Dict[int, tuple] = {}

    for row_idx in range(header_start_row, data_end_row + 1):
        max_lines_in_row = 1  # At least 1 line per row

        for col_idx in range(1, num_columns + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell_value = cell.value

            if cell_value is None:
                continue

            # Skip formula cells — their string representation (=SUM...)
            # doesn't reflect displayed width
            if isinstance(cell_value, str) and cell_value.startswith('='):
                continue

            # Track if this cell is text (not a number/formula)
            is_text = isinstance(cell_value, str)

            # Use formatted display text instead of raw str()
            text = _estimate_display_text(cell_value, cell.number_format)

            # --- Width: track longest display text per column ---
            # For multi-line cells, use the longest single line
            lines = text.split('\n')
            longest_line_len = max(len(line) for line in lines) if lines else 0

            current_max, current_is_text = col_max_info.get(col_idx, (0, False))
            if longest_line_len > current_max:
                col_max_info[col_idx] = (longest_line_len, is_text)

            # --- Height: track max line count per row ---
            line_count = len(lines)
            if line_count > max_lines_in_row:
                max_lines_in_row = line_count

        # Apply row height if there are line breaks (more than 1 line)
        if max_lines_in_row > 1:
            computed_height = max_lines_in_row * line_height
            worksheet.row_dimensions[row_idx].height = computed_height
            logger.debug(f"auto_fit row {row_idx}: height={computed_height} ({max_lines_in_row} lines)")

    # Apply column widths — text gets 1.2x multiplier (wider font chars), numbers stay 1:1
    for col_idx, (max_len, is_text) in col_max_info.items():
        if is_text:
            # Text chars are wider than digits in proportional fonts
            computed_width = max_len * 1.2 + padding + 5
        else:
            computed_width = max_len + padding
        # Clamp between min and max
        computed_width = max(min_width, min(computed_width, max_width))
        col_letter = get_column_letter(col_idx)
        worksheet.column_dimensions[col_letter].width = computed_width
        logger.debug(f"auto_fit col {col_letter}: width={computed_width:.1f} (max_len={max_len})")

    logger.info(f"auto_fit_dimensions: applied {len(col_max_info)} column widths")


def calculate_header_dimensions(header_layout: List[Dict[str, Any]]) -> Tuple[int, int]:
    """
    Calculates the total number of rows and columns a header will occupy.
    """
    if not header_layout:
        return (0, 0)
    num_rows = max(cell.get('row', 0) + cell.get('rowspan', 1) for cell in header_layout)
    num_cols = max(cell.get('col', 0) + cell.get('colspan', 1) for cell in header_layout)
    return (num_rows, num_cols)

def merge_contiguous_cells_by_id(
    worksheet: Worksheet,
    start_row: int,
    end_row: int,
    col_id_to_merge: str,
    column_id_map: Dict[str, int]
):
    """
    Finds and merges contiguous vertical cells within a column that have the same value.
    This is called AFTER all data has been written to the sheet.
    """
    col_idx = column_id_map.get(col_id_to_merge)
    if not col_idx or start_row >= end_row:
        return

    current_merge_start_row = start_row
    value_to_match = worksheet.cell(row=start_row, column=col_idx).value

    for row_idx in range(start_row + 1, end_row + 2):
        cell_value = worksheet.cell(row=row_idx, column=col_idx).value if row_idx <= end_row else object()
        if cell_value != value_to_match:
            if row_idx - 1 > current_merge_start_row:
                if value_to_match is not None and str(value_to_match).strip():
                    try:
                        worksheet.merge_cells(
                            start_row=current_merge_start_row,
                            start_column=col_idx,
                            end_row=row_idx - 1, end_column=col_idx
                        )
                    except Exception as e:
                        logger.error(f"Could not merge cells for ID {col_id_to_merge} from row {current_merge_start_row} to {row_idx - 1}. Error: {e}")
            current_merge_start_row = row_idx
            if row_idx <= end_row:
                value_to_match = cell_value


                value_to_match = cell_value