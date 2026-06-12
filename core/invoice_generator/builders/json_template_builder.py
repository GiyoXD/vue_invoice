import logging
import copy
from typing import List, Dict, Any, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from core.blueprint_generator.internal.scanner.models import TemplateLayout

# Utils

logger = logging.getLogger(__name__)

class JsonTemplateStateBuilder:
    """
    JsonTemplateStateBuilder: Reconstructs Excel template state from JSON configuration.

    This builder is responsible for "hydrating" a template state (headers, footers, styles, merges, dimensions)
    directly from a JSON dictionary, bypassing the need to open and scan a physical .xlsx template file.
    It is a drop-in replacement for the scanning logic found in `TemplateStateBuilder`, designed to work
    with the `layout_template.json` structure.

    Key Responsibilities:
    1.  **Parsing**: Converts coordinate-based JSON data (e.g., "A1": {...}) into row-based grid structures.
    2.  **Style Reconstruction**: Re-creates OpenPyXL style objects (Font, Border, Fill, Alignment) from their JSON representations.
    3.  **State Management**: Maintains separate states for the Header (top of sheet) and Footer (bottom of sheet).
    4.  **Restoration**: Provides methods (`restore_header_only`, `restore_template_footer`) to write this state onto a new worksheet.

    Usage:
        layout_data = loaded_json['template_layout']['Invoice']
        builder = JsonTemplateStateBuilder(layout_data)
        
        # Later, apply to a target worksheet:
        builder.restore_header_only(target_ws)
        builder.restore_template_footer(target_ws, footer_start_row=50)
    """
    
    DEBUG = False

    def __init__(self, sheet_layout_data: Dict[str, Any], debug: bool = False):
        """
        Initialize and populate state from JSON data.
        
        Args:
            sheet_layout_data: The dictionary for a specific sheet from the layout_template.json
                               (e.g., loaded_json['template_layout']['Invoice'])
            debug: Enable debug printing
        """
        self.layout_data = sheet_layout_data
        self.debug = debug or self.DEBUG
        
        # DEBUG INPUT
        logger.debug(f"[JsonTemplateStateBuilder] __init__ INPUT: sheet_layout_data keys={list(sheet_layout_data.keys()) if sheet_layout_data else 'None'}")
        
        # State structures (same as TemplateStateBuilder)
        self.header_state: List[List[Dict[str, Any]]] = []
        self.footer_state: List[List[Dict[str, Any]]] = []
        self.header_merged_cells: List[str] = []
        self.footer_merged_cells: List[str] = []
        self.row_heights: Dict[int, float] = {}
        self.column_widths: Dict[int, float] = {}
        
        # Relative Footer State (0-indexed)
        # These store the footer structure decoupled from absolute template coordinates.
        # Row 0 = The first row of the footer block.
        self.relative_footer_row_heights: Dict[int, float] = {}
        self.relative_footer_merges: List[tuple] = [] # List of (min_col, min_row, max_col, max_row)
        
        # Row tracking
        self.template_footer_start_row: int = -1
        self.template_footer_end_row: int = -1
        self.header_end_row: int = -1
        
        # Dimensions
        self.min_row = 1
        self.max_row = 1
        self.min_col = 1
        self.max_col = 1
        
        # Column mapping for shifting
        self.column_mapping: Dict[int, int] = {}

        # Parse the JSON data immediately
        self._parse_layout_data()

    def set_column_mapping(self, mapping: Dict[int, int]):
        """Set the column mapping for restoration (same as TemplateStateBuilder)."""
        self.column_mapping = mapping

    def _get_mapped_column(self, template_col: int) -> int:
        """Get output column index (same as TemplateStateBuilder)."""
        if not self.column_mapping:
            return template_col
        return self.column_mapping.get(template_col, template_col)

    def _parse_layout_data(self):
        """
        Parses the raw JSON layout data into internal state structures.
        """
        logger.info("[JsonTemplateStateBuilder] Parsing Layout Data")
        
        # Initialize TemplateLayout from dict
        self.layout_obj = TemplateLayout.from_dict(self.layout_data)
        
        # 1. Parse Dimensions and Basic Props
        for col_letter, width in self.layout_obj.col_widths.items():
            self.column_widths[column_index_from_string(col_letter)] = width

        # 2. Parse Header State and find header boundaries
        header_end_row = 0
        max_col = 1
        for row in self.layout_obj.header_rows:
            r_idx = row.relative_index + 1
            header_end_row = max(header_end_row, r_idx)
            for cell in row.cells:
                max_col = max(max_col, cell.col_index)
                if cell.merge:
                    max_col = max(max_col, cell.merge.max_col)
        
        self.header_end_row = header_end_row
        self.max_col = max_col

        # Populate row heights for header
        for row in self.layout_obj.header_rows:
            r_idx = row.relative_index + 1
            if row.height is not None:
                self.row_heights[r_idx] = row.height

        # Populate header merged cells list
        self.header_merged_cells = []
        for row in self.layout_obj.header_rows:
            row_idx = row.relative_index + 1
            for cell in row.cells:
                if cell.merge:
                    range_str = f"{get_column_letter(cell.merge.min_col)}{row_idx}:{get_column_letter(cell.merge.max_col)}{row_idx + cell.merge.row_span - 1}"
                    self.header_merged_cells.append(range_str)

        # Build header_state grid
        header_state = []
        for r in range(1, header_end_row + 1):
            row_data = []
            for c in range(1, max_col + 1):
                row_data.append({
                    'value': None,
                    'font': None,
                    'fill': None,
                    'border': None,
                    'alignment': None,
                    'number_format': 'General'
                })
            header_state.append(row_data)

        for row in self.layout_obj.header_rows:
            r_idx = row.relative_index + 1
            for cell in row.cells:
                g_r = r_idx - 1
                g_c = cell.col_index - 1
                
                style_dict = cell.style.to_dict() if cell.style else {}
                
                cell_info = {
                    'value': cell.value,
                    'font': self._create_font(style_dict.get('font')),
                    'fill': self._create_fill(style_dict.get('fill')),
                    'border': self._create_border(style_dict.get('border')),
                    'alignment': self._create_alignment(style_dict.get('alignment')),
                    'number_format': style_dict.get('number_format', 'General')
                }
                
                if 0 <= g_r < len(header_state) and 0 <= g_c < len(header_state[g_r]):
                    header_state[g_r][g_c] = cell_info
        
        self.header_state = header_state

        # 3. Parse Template Footer State
        self.template_footer_rows = []
        for row in self.layout_obj.footer_rows:
            row_dict = {
                "relative_index": row.relative_index,
                "height": row.height,
                "cells": [],
                "merges": []
            }
            for cell in row.cells:
                c_dict = {"col_index": cell.col_index}
                if cell.value is not None:
                    c_dict["value"] = cell.value
                if cell.style:
                    c_dict["style"] = cell.style.to_dict()
                row_dict["cells"].append(c_dict)
                
                if cell.merge:
                    row_dict["merges"].append({
                        "min_col": cell.merge.min_col,
                        "max_col": cell.merge.max_col,
                        "row_span": cell.merge.row_span,
                        "value": cell.merge.value
                    })
            self.template_footer_rows.append(row_dict)

        if self.template_footer_rows:
            if self.header_end_row <= 0:
                logger.error(
                    "[JsonTemplateStateBuilder] header_end_row is 0 or negative — "
                    "cannot safely place footer. template_footer_start_row set to -1. "
                    "Check that header_content is non-empty in the layout JSON."
                )
                self.template_footer_start_row = -1
            else:
                self.template_footer_start_row = self.header_end_row + 1
            max_rel_idx = max((r.relative_index for r in self.layout_obj.footer_rows), default=-1)
            self.template_footer_end_row = (self.template_footer_start_row + max_rel_idx) if max_rel_idx >= 0 else -1
            
            # Update max_col based on new footer cells
            for row in self.layout_obj.footer_rows:
                for cell in row.cells:
                    self.max_col = max(self.max_col, cell.col_index)
                    if cell.merge:
                        self.max_col = max(self.max_col, cell.merge.max_col)
        else:
            self.template_footer_start_row = -1
            self.template_footer_end_row = -1

        # Update max_col
        if self.column_widths:
            self.max_col = max(self.max_col, max(self.column_widths.keys()))
        
        # Update max_row
        if self.template_footer_end_row > 0:
            self.max_row = self.template_footer_end_row
        elif self.header_end_row > 0:
            self.max_row = self.header_end_row



    def _parse_color(self, c_val) -> Optional[str]:
        if not c_val: return None
        if isinstance(c_val, dict):
            if 'rgb' in c_val:
                c_val = c_val['rgb']
            else:
                return None
        
        if isinstance(c_val, str):
            if c_val.startswith("theme-"):
                return None
            if c_val.lower() == 'auto' or c_val == '00000000':
                return None
            
            # Clean and normalize hex
            c_val = c_val.lstrip('#')
            # Handle standard RGB (e.g. openpyxl might spit out 6 chars or users might enter it)
            if len(c_val) == 6:
                return 'FF' + c_val
            if len(c_val) == 8:
                # Basic check for hex characters
                try:
                    int(c_val, 16)
                    return c_val
                except ValueError:
                    return None
        return None

    def _create_font(self, d: Dict) -> Optional[Font]:
        if not d: return None
        # Handle color dict/str safely
        color = self._parse_color(d.get('color'))
             
        return Font(
            name=d.get('name'),
            size=d.get('size'),
            bold=d.get('bold'),
            italic=d.get('italic'),
            strike=d.get('strike'),
            underline=d.get('underline'),
            color=color
        )
        
    def _create_fill(self, d: Dict) -> Optional[PatternFill]:
        if not d: return None
        if not d.get('type') or d.get('type') == 'none': return None
        
        fgColor = self._parse_color(d.get('color'))
        if not fgColor: return None
        
        return PatternFill(
            fill_type=d.get('type'),
            start_color=fgColor,
            end_color=fgColor # Simple solid fill assumption
        )
        
    def _create_border(self, d: Dict) -> Optional[Border]:
        if not d: return None
        def _side(s_data):
            if not s_data: return None
            # s_data might be simple style string or dict? 
            # Review sanitizer: "left": cell.border.left.style
            # It saves just the style string (e.g. 'thin', 'medium')
            return Side(style=s_data) if s_data else None

        return Border(
            left=_side(d.get('left')),
            right=_side(d.get('right')),
            top=_side(d.get('top')),
            bottom=_side(d.get('bottom'))
        )

    def _create_alignment(self, d: Dict) -> Optional[Alignment]:
        if not d: return None
        return Alignment(
            horizontal=d.get('horizontal'),
            vertical=d.get('vertical'),
            text_rotation=d.get('text_rotation', 0),
            wrap_text=d.get('wrap_text'),
            shrink_to_fit=d.get('shrink_to_fit'),
            indent=d.get('indent', 0)
        )

    # --- Restoration Logic (Mirrors TemplateStateBuilder) ---
    # We copy this verbatim from TemplateStateBuilder to allow safe refactor later.
    
    def restore_header_only(self, target_worksheet: Worksheet, actual_num_cols: int = None, mode: str = "standard"):
        """
        Restores ONLY the header to a new clean worksheet.
        
        Args:
            target_worksheet: The worksheet to write header content onto.
            actual_num_cols: Optional column count to limit restoration.
            mode: Generation mode ('standard', 'daf', 'custom'). Used to resolve
                  mode-dependent cell values in header_content.
        """
        logger.info(f"[JsonTemplateStateBuilder] Restoring Header to '{target_worksheet.title}' (mode={mode})")
        logger.debug(f"[JsonTemplateStateBuilder] restore_header_only INPUT: target_worksheet={target_worksheet.title}, actual_num_cols={actual_num_cols}, mode={mode}")

        template_num_cols = self.max_col
        target_num_cols = actual_num_cols if actual_num_cols else template_num_cols
        
        # Restore header cell values and formatting
        for row_idx, row_data in enumerate(self.header_state):
            # For header, we start at min_row (usually 1)
            actual_row = row_idx + self.min_row
            
            for col_idx, cell_info in enumerate(row_data):
                template_col = col_idx + self.min_col
                output_col = self._get_mapped_column(template_col)
                
                if output_col is None:
                    continue # Skip removed columns (simple version of logic)
                
                target_cell = target_worksheet.cell(row=actual_row, column=output_col)
                self._write_cell(target_cell, cell_info, mode=mode)
                
        # Restore header merges
        for merge_str in self.header_merged_cells:
            self._apply_merge(target_worksheet, merge_str)
            
        # Restore dimensions
        for r_idx in range(self.min_row, self.header_end_row + 1):
             if r_idx in self.row_heights:
                 target_worksheet.row_dimensions[r_idx].height = self.row_heights[r_idx]
                 
        for c_idx, w in self.column_widths.items():
            target_worksheet.column_dimensions[get_column_letter(c_idx)].width = w

    def restore_template_footer(self, target_worksheet: Worksheet, footer_start_row: int, actual_num_cols: int = None, mode: str = "standard"):
        """
        Restores the template footer content onto the target worksheet at a specific starting row.
        """
        logger.info(f"[JsonTemplateStateBuilder] Restoring Footer to '{target_worksheet.title}' at row {footer_start_row} (mode={mode})")

        # --- NEW GRID-ROW FORMAT ---
        if hasattr(self, 'template_footer_rows') and self.template_footer_rows is not None:
            if not self.template_footer_rows:
                logger.warning(f"[JsonTemplateStateBuilder] Template footer rows is empty for '{target_worksheet.title}'.")
                return
                
            skip_count = 0
            for row_dict in self.template_footer_rows:
                # Backward-compat: old JSONs may still have is_dynamic_footer=True on the TOTAL row
                # (relative_index=0). New JSONs never include it — sanitizer now starts capture at
                # table_footer_row+1, so skip_count stays 0 for new JSONs and math is identical.
                if row_dict.get('is_dynamic_footer'):
                    skip_count += 1
                    continue

                rel_idx = row_dict.get('relative_index', 0)
                actual_row = footer_start_row + rel_idx - skip_count
                
                # 1. Restore Row Height
                h = row_dict.get('height')
                if h is not None:
                    target_worksheet.row_dimensions[actual_row].height = h
                    
                # 2. Restore Cells (Values & Styles)
                for cell_dict in row_dict.get('cells', []):
                    template_col = cell_dict.get('col_index')
                    output_col = self._get_mapped_column(template_col)
                    
                    if output_col is None: continue
                    
                    target_cell = target_worksheet.cell(row=actual_row, column=output_col)
                    
                    val = cell_dict.get('value')
                    if val is not None:
                        resolved = self._resolve_mode_value(val, mode)
                        if resolved is not None:
                            target_cell.value = resolved
                        
                    style_dict = cell_dict.get('style')
                    if style_dict:
                        font = self._create_font(style_dict.get('font'))
                        fill = self._create_fill(style_dict.get('fill'))
                        border = self._create_border(style_dict.get('border'))
                        align = self._create_alignment(style_dict.get('alignment'))
                        num_fmt = style_dict.get('number_format', 'General')
                        
                        if font: target_cell.font = copy.copy(font)
                        if fill: target_cell.fill = copy.copy(fill)
                        if border: target_cell.border = copy.copy(border)
                        if align: target_cell.alignment = copy.copy(align)
                        if num_fmt: target_cell.number_format = num_fmt
                        
                # 3. Restore Merges
                for m_dict in row_dict.get('merges', []):
                    min_col = m_dict.get('min_col')
                    max_col = m_dict.get('max_col')
                    row_span = m_dict.get('row_span', 1)
                    
                    mapped_min_col = self._get_mapped_column(min_col)
                    mapped_max_col = self._get_mapped_column(max_col)
                    
                    if mapped_min_col and mapped_max_col:
                        new_range = f"{get_column_letter(mapped_min_col)}{actual_row}:{get_column_letter(mapped_max_col)}{actual_row + row_span - 1}"
                        try:
                            target_worksheet.merge_cells(new_range)
                        except ValueError:
                            logger.warning(f"[JsonTemplateStateBuilder] Skipped overlapping merge {new_range} on '{target_worksheet.title}'.")
            return
            


    @staticmethod
    def _resolve_mode_value(raw_value, mode: str = "standard"):
        """
        Resolves a cell value that may be mode-dependent.
        
        Resolution priority (highest to lowest):
            1. Exact mode-specific override (e.g., 'daf' key when mode='daf')
            2. 'standard' — the UNIVERSAL base value that applies to ALL modes
            3. 'default' — the original/fallback value
        
        This means 'standard' is NOT mode-specific; it acts as a universal
        override that applies to standard, custom, DAF, and any future mode.
        Only a mode-specific key (e.g., 'daf') can take priority over it.
        
        Args:
            raw_value: The raw value from header_content (str, number, or dict).
            mode: The active generation mode ('standard', 'daf', 'custom').
        
        Returns:
            The resolved scalar value.
        """
        if isinstance(raw_value, dict):
            # 1. Exact mode-specific override (e.g., 'daf' key when mode='daf')
            #    This does NOT match 'standard' key when mode='standard' — that's
            #    handled by step 2 as the universal base.
            if mode != "standard" and mode in raw_value:
                return raw_value[mode]
            
            # 2. 'standard' = universal base value (applies to ALL modes)
            if "standard" in raw_value:
                return raw_value["standard"]
                
            # 3. 'default' fallback (original template value).
            # Return None if no 'default' key — means "no override for this mode,
            # keep whatever the template had." Do NOT leak another mode's value (e.g.
            # a daf-only dict must not apply its value when mode='standard').
            return raw_value.get('default', None)
            
        return raw_value

    def _write_cell(self, cell, info, mode: str = "standard"):
        """
        Writes a single cell's state (value and styles) to an OpenPyXL cell object.
        
        Args:
            cell: The target OpenPyXL Cell object.
            info: A dictionary containing 'value', 'font', 'fill', 'border', 'alignment', 'number_format'.
            mode: Generation mode for resolving mode-dependent values.
        """
        if info['value'] is not None:
            resolved = self._resolve_mode_value(info['value'], mode)
            if resolved is not None:
                cell.value = resolved
        if info['font']: cell.font = copy.copy(info['font'])
        if info['fill']: cell.fill = copy.copy(info['fill'])
        if info['border']: cell.border = copy.copy(info['border'])
        if info['alignment']: cell.alignment = copy.copy(info['alignment'])
        if info['number_format']: cell.number_format = info['number_format']

    def _apply_merge(self, ws, merge_data, start_row_offset=0):
        """
        Applies a merge range to the worksheet.
        
        Args:
            ws: The target worksheet.
            merge_data: Either a string "A1:B2" (absolute) OR a tuple (min_col, min_r, max_col, max_r) (relative).
            start_row_offset: Offset to add to row indices (typically the footer start row).
        """
        min_col, min_row, max_col, max_row = 0, 0, 0, 0
        
        # Determine input type
        if isinstance(merge_data, str):
            # Classic string parsing (used by Header) - Absolute
            from openpyxl.utils.cell import range_boundaries
            min_col, min_row, max_col, max_row = range_boundaries(merge_data)
            # No offset usually needed for absolute strings, unless shifted?
            # Existing logic was confusing. For Header, we use it as-is.
        elif isinstance(merge_data, tuple) or isinstance(merge_data, list):
            # Relative tuple (used by Footer) - (col, rel_row, col, rel_row)
            min_col, rel_min, max_col, rel_max = merge_data
            min_row = rel_min + start_row_offset
            max_row = rel_max + start_row_offset
            
        # Apply column mapping
        mapped_min_col = self._get_mapped_column(min_col)
        mapped_max_col = self._get_mapped_column(max_col)
        
        if mapped_min_col and mapped_max_col:
            new_range = f"{get_column_letter(mapped_min_col)}{min_row}:{get_column_letter(mapped_max_col)}{max_row}"
            try:
                ws.merge_cells(new_range)
            except ValueError:
                # Overlapping merges can cause this
                pass


