# invoice_generator/styling/style_applier.py
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Side, Font
from typing import Dict, Any, Optional, List, Tuple

# Import centralized style constants
from .style_config import (
    THIN_BORDER, NO_BORDER, CENTER_ALIGNMENT, LEFT_ALIGNMENT, BOLD_FONT, SIDE_BORDER
)

# --- Constants for Number Formats ---
FORMAT_GENERAL = 'General'
FORMAT_TEXT = '@'
FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0'
FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00'

from .models import StylingConfigModel

def apply_cell_style(cell: Worksheet.cell, styling_config: StylingConfigModel, context: dict):
    """
    Applies all styles to a single cell, including fonts, alignments,
    and complex conditional borders, based on its context.
    """
    import logging
    logger = logging.getLogger(__name__)
    
    # --- Get Context ---
    col_id = context.get("col_id")
    col_idx = context.get("col_idx")
    static_col_idx = context.get("static_col_idx")
    is_pre_footer = context.get("is_pre_footer", False)
    is_static_row = context.get("is_static_row", False)
    is_header = context.get("is_header", False)
    DAF_mode = context.get("DAF_mode", False)
    
    # Log what context we received
    logger.debug(f"apply_cell_style: col_id={col_id}, col_idx={col_idx}, is_header={is_header}, is_static_row={is_static_row}")
    
    if not styling_config:
        logger.warning(f"apply_cell_style: NO styling_config provided (col_id={col_id}, col_idx={col_idx})")
        return

    # Handle static rows first
    if is_static_row:
        cell.alignment = CENTER_ALIGNMENT
        cell.border = NO_BORDER
        if isinstance(cell.value, (int, float)):
            cell.number_format = FORMAT_NUMBER_COMMA_SEPARATED2 if isinstance(cell.value, float) else FORMAT_NUMBER_COMMA_SEPARATED1
        else:
            cell.number_format = FORMAT_TEXT
        return
        
    if is_header:
        cell.border = THIN_BORDER
        return

    # --- 1. Apply Font, Alignment, and Number Formats ---
    if col_id and styling_config:
        col_specific_style = styling_config.columnIdStyles.get(col_id) if styling_config.columnIdStyles else None
        
        if not col_specific_style:
            logger.debug(f"No column-specific style for col_id={col_id}")
            
        # Check if we have ANY font source
        has_col_font = col_specific_style and col_specific_style.font
        has_default_font = styling_config.defaultFont
        
        if not has_col_font and not has_default_font:
            logger.warning(f"NO font available for col_id={col_id} - neither column-specific nor default font exists")
        
        # Check if we have ANY alignment source
        has_col_alignment = col_specific_style and col_specific_style.alignment
        has_default_alignment = styling_config.defaultAlignment
        
        if not has_col_alignment and not has_default_alignment:
            logger.warning(f"NO alignment available for col_id={col_id} - neither column-specific nor default alignment exists")
        
        if col_specific_style:
            if col_specific_style.font:
                cell.font = Font(**col_specific_style.font.model_dump(exclude_none=True))
            elif styling_config.defaultFont:
                cell.font = Font(**styling_config.defaultFont.model_dump(exclude_none=True))

            if col_specific_style.alignment:
                cell.alignment = Alignment(**col_specific_style.alignment.model_dump(exclude_none=True))
            elif styling_config.defaultAlignment:
                cell.alignment = Alignment(**styling_config.defaultAlignment.model_dump(exclude_none=True))

            # --- Apply Number Format ---
            number_format = col_specific_style.numberFormat
            
            # PCS always uses config format, never forced format
            if col_id in ['col_pcs', 'col_qty_pcs']:
                if number_format and cell.number_format != FORMAT_TEXT:
                    cell.number_format = number_format
            else:
                # Non-PCS columns follow DAF mode logic
                if number_format and cell.number_format != FORMAT_TEXT and not DAF_mode:
                    cell.number_format = number_format
                elif number_format and cell.number_format != FORMAT_TEXT and DAF_mode:
                    cell.number_format = FORMAT_NUMBER_COMMA_SEPARATED2
                elif cell.number_format != FORMAT_TEXT and (cell.number_format == FORMAT_GENERAL or cell.number_format is None):
                    if isinstance(cell.value, float): cell.number_format = FORMAT_NUMBER_COMMA_SEPARATED2
                    elif isinstance(cell.value, int): cell.number_format = FORMAT_NUMBER_COMMA_SEPARATED1

    # --- 2. Apply Conditional Borders ---
    # Special handling for the pre-footer row
    if is_pre_footer:
        if col_idx == static_col_idx:
            cell.border = SIDE_BORDER
        else:
            cell.border = THIN_BORDER
        return

    # UPDATED: Simplified logic for main data rows
    if col_idx == static_col_idx:
        # The static column ONLY ever gets side borders.
        cell.border = SIDE_BORDER
    elif col_idx: 
        # All other columns get a full grid.
        cell.border = THIN_BORDER


def apply_row_heights(worksheet: Worksheet, sheet_styling_config: Optional[StylingConfigModel], header_info: Optional[Dict[str, Any]] = None, data_row_indices: Optional[List[int]] = None, footer_row_index: Optional[int] = None, row_after_header_idx: int = -1, row_before_footer_idx: int = -1):
    import logging
    logging.debug(f"sheet_styling_config: {sheet_styling_config}")
    if not sheet_styling_config or not sheet_styling_config.rowHeights:
        return
    row_heights_cfg = sheet_styling_config.rowHeights

    actual_header_height = None # Store the applied header height

    def set_height(r_idx, height_val, desc): # Helper function
        nonlocal actual_header_height # Ensure actual_header_height is modified
        if r_idx <= 0: return
        try:
            h_val = float(height_val)
            if h_val > 0:
                worksheet.row_dimensions[r_idx].height = h_val
                import logging
                logging.debug(f"Setting row {r_idx} height to {h_val}")
                if desc == "header": # Store the height applied to the header
                    actual_header_height = h_val
            else: pass # Ignore non-positive heights
        except (ValueError, TypeError): pass # Ignore invalid height values
        except Exception as height_err: pass # Log other errors?

    # Apply Heights Based on Config
    header_height = row_heights_cfg.get("header")
    if header_height is not None and header_info:
        h_start = header_info.get('first_row_index', -1); h_end = header_info.get('second_row_index', -1)
        if h_start > 0 and h_end >= h_start:
            for r in range(h_start, h_end + 1): set_height(r, header_height, "header")

    after_header_height = row_heights_cfg.get("after_header")
    if after_header_height is not None and row_after_header_idx > 0: set_height(row_after_header_idx, after_header_height, "after_header")
    data_default_height = row_heights_cfg.get("data_default")
    if data_default_height is not None and data_row_indices:
        for r in data_row_indices: set_height(r, data_default_height, "data_default")
    before_footer_height = row_heights_cfg.get("before_footer")
    if before_footer_height is not None and row_before_footer_idx > 0: set_height(row_before_footer_idx, before_footer_height, "before_footer")

    # --- Footer Height Logic ---
    footer_height_config = row_heights_cfg.get("footer")
    match_header_height_flag = row_heights_cfg.get("footer_matches_header_height", True) # Default to True

    final_footer_height = None
    if match_header_height_flag and actual_header_height is not None:
        final_footer_height = actual_header_height # Use header height if flag is true and header height was set
    elif footer_height_config is not None:
        final_footer_height = footer_height_config # Otherwise, use specific footer height if defined

    if final_footer_height is not None and footer_row_index > 0:
        set_height(footer_row_index, final_footer_height, "footer")
    # --- End Footer Height Logic ---

    specific_heights = row_heights_cfg.get("specific_rows")
    if isinstance(specific_heights, dict):
        for row_str, height_val in specific_heights.items():
            try: row_num = int(row_str); set_height(row_num, height_val, f"specific_row_{row_num}")
            except ValueError: pass # Ignore invalid row numbers

def apply_header_style(cell: Worksheet.cell, styling_config: StylingConfigModel):
    """
    Applies styling to a header cell, using config values with fallbacks.
    """
    import logging
    logger = logging.getLogger(__name__)
    
    if not styling_config:
        logger.warning(f"apply_header_style called with NO styling_config - cannot apply header styles")
        return
    
    if not styling_config.headerFont:
        logger.warning(f"apply_header_style: styling_config has NO headerFont - header will have no font styling")
        logger.warning(f"  styling_config attributes: {list(styling_config.model_dump().keys())}")
    else:
        font_dict = styling_config.headerFont.model_dump(exclude_none=True)
        logger.debug(f"Applying header font: {font_dict}")
        
        # Check for missing critical font properties
        if not font_dict.get('name'):
            logger.warning(f"headerFont missing 'name' property")
        if not font_dict.get('size'):
            logger.warning(f"headerFont missing 'size' property")
        
        # Create Font object explicitly
        effective_header_font = Font(
            name=font_dict.get('name'),
            size=font_dict.get('size'),
            bold=font_dict.get('bold'),
            italic=font_dict.get('italic'),
            color=font_dict.get('color'),
            family=2,
            scheme='minor'
        )
        cell.font = effective_header_font
        logger.debug(f"Applied font: name={effective_header_font.name}, size={effective_header_font.size}, bold={effective_header_font.bold}")
    
    if not styling_config.headerAlignment:
        logger.warning(f"apply_header_style: styling_config has NO headerAlignment - header will have no alignment")
    else:
        effective_header_align = Alignment(**styling_config.headerAlignment.model_dump(exclude_none=True))
        cell.alignment = effective_header_align
        logger.debug(f"Applied alignment: {effective_header_align.horizontal}, {effective_header_align.vertical}")
