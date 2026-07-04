import logging
import copy
from typing import List, Dict, Any, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from core.blueprint_generator.internal.scanner.models import TemplateLayout, ZoneBoundaries
from core.models.cell import UnitRow, UnitCell, TemplateMerge

# Utils

logger = logging.getLogger(__name__)

class JsonTemplateStateBuilder:
    """
    JsonTemplateStateBuilder: Reconstructs Excel template state from JSON configuration.

    This builder reconstructs template visual state (headers, footers, styles, merges, dimensions)
    from a JSON dictionary using TemplateLayout model objects (UnitRow, UnitCell, CellStyle, TemplateMerge).
    
    It uses a single unified `restore_rows()` method to write any list of UnitRow objects
    to a worksheet at a given position — no separate code paths for header vs footer.

    Key Responsibilities:
    1.  **Parsing**: Initializes TemplateLayout from JSON and calculates boundaries.
    2.  **Style Reconstruction**: Re-creates OpenPyXL style objects (Font, Border, Fill, Alignment) from model dicts.
    3.  **Restoration**: Uses `restore_rows()` to write UnitRow model objects directly to a worksheet.

    Usage:
        layout_data = loaded_json['template_layout']['Invoice']
        builder = JsonTemplateStateBuilder(layout_data)
        
        # Later, apply to a target worksheet:
        builder.restore_header_only(target_ws)
        builder.restore_template_footer(target_ws, footer_start_row=50)
    """
    
    def __init__(self, sheet_layout_data: Dict[str, Any]):
        """
        Initialize and populate state from JSON data.
        
        Args:
            sheet_layout_data: The dictionary for a specific sheet from the layout_template.json
                               (e.g., loaded_json['template_layout']['Invoice'])
        """
        self.layout_data = sheet_layout_data
        
        # DEBUG INPUT
        logger.debug(f"[JsonTemplateStateBuilder] __init__ INPUT: sheet_layout_data keys={list(sheet_layout_data.keys()) if sheet_layout_data else 'None'}")
        
        # Boundaries
        self.boundaries: Optional[ZoneBoundaries] = None
        
        # Dimensions
        self.min_row = 1
        self.max_row = 1
        self.min_col = 1
        self.max_col = 1
        
        # Parse the JSON data immediately
        self._parse_layout_data()

    def _parse_layout_data(self):
        """
        Parses the JSON layout data into TemplateLayout model and calculates boundaries.
        All row/cell/style data lives in self.layout_obj — no intermediate dicts.
        """
        logger.info("[JsonTemplateStateBuilder] Parsing Layout Data")
        
        # Initialize TemplateLayout from dict
        self.layout_obj = TemplateLayout.from_dict(self.layout_data)
        
        # 2. Calculate header boundaries
        header_end_row = 0
        max_col = 1
        for row in self.layout_obj.header_rows:
            r_idx = row.relative_index + 1
            header_end_row = max(header_end_row, r_idx)
            for cell in row.cells:
                max_col = max(max_col, cell.col_index)
                if cell.merge:
                    max_col = max(max_col, cell.merge.max_col)
        self.max_col = max_col

        # 2. Calculate footer boundaries
        template_footer_start_row = -1
        template_footer_end_row = -1
        if self.layout_obj.footer_rows:
            if header_end_row <= 0:
                logger.error(
                    "[JsonTemplateStateBuilder] header_end_row is 0 or negative — "
                    "cannot safely place footer. template_footer_start_row set to -1. "
                    "Check that header_content is non-empty in the layout JSON."
                )
            else:
                template_footer_start_row = header_end_row + 1
            max_rel_idx = max((r.relative_index for r in self.layout_obj.footer_rows), default=-1)
            template_footer_end_row = (template_footer_start_row + max_rel_idx) if max_rel_idx >= 0 else -1
            
            # Update max_col based on footer cells
            for row in self.layout_obj.footer_rows:
                for cell in row.cells:
                    self.max_col = max(self.max_col, cell.col_index)
                    if cell.merge:
                        self.max_col = max(self.max_col, cell.merge.max_col)
        

        
        if template_footer_end_row > 0:
            self.max_row = template_footer_end_row
        elif header_end_row > 0:
            self.max_row = header_end_row
        else:
            self.max_row = 1

        self.boundaries = ZoneBoundaries(
            header_row=header_end_row + 1,
            data_start_row=header_end_row + 2,
            footer_row=template_footer_start_row - 1 if template_footer_start_row > 0 else None,
            max_col=self.max_col
        )

    # --- Boundaries Compatibility Properties ---

    @property
    def header_end_row(self) -> int:
        return self.boundaries.header_row - 1 if self.boundaries else -1

    @property
    def template_footer_start_row(self) -> int:
        if not self.boundaries or self.boundaries.footer_row is None:
            return -1
        return self.boundaries.footer_row + 1

    @property
    def template_footer_end_row(self) -> int:
        if not self.boundaries or self.boundaries.footer_row is None:
            return -1
        max_rel_idx = max((r.relative_index for r in self.layout_obj.footer_rows), default=-1)
        return (self.template_footer_start_row + max_rel_idx) if max_rel_idx >= 0 else -1

    # --- Style Helpers ---

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

    # --- Core Restoration Logic ---

    def restore_rows(self, ws: Worksheet, rows: List[UnitRow],
                     start_row: int, mode: str = "standard",
                     layout_state: Optional[Any] = None):
        """
        Unified restoration: writes a list of UnitRow model objects to a worksheet.
        
        Works for both header and footer — the caller decides what rows and where.
        Iterates model objects directly (UnitRow → UnitCell → CellStyle/TemplateMerge),
        no intermediate dictionaries.
        
        Args:
            ws: Target worksheet to write onto.
            rows: List of UnitRow model objects (from layout_obj.header_rows or footer_rows).
            start_row: The worksheet row to map relative_index=0 onto.
            mode: Generation mode for resolving mode-dependent values ('standard', 'daf', 'custom').
            layout_state: Optional layout state tracking occupied/merged cells.
        """
        for row in rows:
            actual_row = start_row + row.relative_index
            
            # Row height
            if row.height is not None:
                ws.row_dimensions[actual_row].height = row.height
            
            # Cells
            for cell in row.cells:
                target = ws.cell(row=actual_row, column=cell.col_index)
                
                # Value (with mode resolution)
                if cell.value is not None:
                    resolved = self._resolve_mode_value(cell.value, mode)
                    if resolved is not None:
                        target.value = resolved
                
                # Style
                if cell.style:
                    sd = cell.style.to_dict()
                    font = self._create_font(sd.get('font'))
                    fill = self._create_fill(sd.get('fill'))
                    border = self._create_border(sd.get('border'))
                    align = self._create_alignment(sd.get('alignment'))
                    num_fmt = sd.get('number_format', 'General')
                    
                    if font: target.font = copy.copy(font)
                    if fill: target.fill = copy.copy(fill)
                    if border: target.border = copy.copy(border)
                    if align: target.alignment = copy.copy(align)
                    if num_fmt: target.number_format = num_fmt
                
                # Merge
                if cell.merge:
                    merge_range_str = (
                        f"{get_column_letter(cell.merge.min_col)}{actual_row}:"
                        f"{get_column_letter(cell.merge.max_col)}"
                        f"{actual_row + cell.merge.row_span - 1}"
                    )
                    
                    if layout_state:
                        # Register in coordinator to track occupied and merged cells
                        layout_state.merge_cells(
                            start_row=actual_row,
                            start_col_id_or_idx=cell.merge.min_col,
                            end_row=actual_row + cell.merge.row_span - 1,
                            end_col_id_or_idx=cell.merge.max_col
                        )
                    else:
                        # Fallback to direct merge without coordinator tracking
                        # Check if cells are already merged to prevent ValueError
                        from openpyxl.utils import range_boundaries
                        min_col, min_row, max_col, max_row = range_boundaries(merge_range_str)
                        
                        overlap = False
                        for existing_range in ws.merged_cells.ranges:
                            # Check intersection
                            if (min_row <= existing_range.max_row and max_row >= existing_range.min_row and
                                min_col <= existing_range.max_col and max_col >= existing_range.min_col):
                                overlap = True
                                break
                                
                        if overlap:
                            logger.warning(f"[JsonTemplateStateBuilder] Skipped overlapping merge {merge_range_str} on '{ws.title}'.")
                        else:
                            try:
                                ws.merge_cells(merge_range_str)
                            except ValueError as e:
                                logger.warning(f"[JsonTemplateStateBuilder] Merge failed {merge_range_str}: {e}")

    # --- Public API (thin wrappers) ---
    
    def restore_header_only(self, target_worksheet: Worksheet, actual_num_cols: int = None, mode: str = "standard", layout_state: Optional[Any] = None, column_index_mapping: Optional[Dict[int, Optional[int]]] = None):
        """
        Restores ONLY the header to a worksheet.
        
        Args:
            target_worksheet: The worksheet to write header content onto.
            actual_num_cols: Optional column count (kept for API compatibility, not used).
            mode: Generation mode ('standard', 'daf', 'custom').
            layout_state: Optional layout state tracking occupied/merged cells.
            column_index_mapping: Optional template to physical column mapping index.
        """
        logger.info(f"[JsonTemplateStateBuilder] Restoring Header to '{target_worksheet.title}' (mode={mode})")
        
        rows = self.layout_obj.header_rows
        if column_index_mapping:
            rows = translate_template_rows(rows, column_index_mapping)
            
        self.restore_rows(target_worksheet, rows, start_row=1, mode=mode, layout_state=layout_state)
        


    def restore_template_footer(self, target_worksheet: Worksheet, footer_start_row: int, actual_num_cols: int = None, mode: str = "standard", layout_state: Optional[Any] = None, column_index_mapping: Optional[Dict[int, Optional[int]]] = None):
        """
        Restores the template footer content at a specific starting row.
        
        Args:
            target_worksheet: The worksheet to write footer content onto.
            footer_start_row: The row number where the footer should begin.
            actual_num_cols: Optional column count (kept for API compatibility, not used).
            mode: Generation mode ('standard', 'daf', 'custom').
            layout_state: Optional layout state tracking occupied/merged cells.
            column_index_mapping: Optional template to physical column mapping index.
        """
        logger.info(f"[JsonTemplateStateBuilder] Restoring Footer to '{target_worksheet.title}' at row {footer_start_row} (mode={mode})")
        
        if not self.layout_obj.footer_rows:
            logger.warning(f"[JsonTemplateStateBuilder] Template footer rows is empty for '{target_worksheet.title}'.")
            return
            
        rows = self.layout_obj.footer_rows
        if column_index_mapping:
            rows = translate_template_rows(rows, column_index_mapping)
        
        self.restore_rows(target_worksheet, rows, start_row=footer_start_row, mode=mode, layout_state=layout_state)

    # --- Value Resolution ---

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


def translate_template_rows(rows: List[UnitRow], column_mapping: Dict[int, Optional[int]]) -> List[UnitRow]:
    """
    Translates template UnitRow cells' col_index using the template-to-physical column mapping.
    Filters out cells that map to None (hidden columns) and updates merge ranges.
    """
    if not column_mapping:
        return rows
        
    translated_rows = []
    for row in rows:
        translated_cells = []
        covered_template_cols = set()
        
        for cell in row.cells:
            if cell.col_index in covered_template_cols:
                continue
                
            if cell.merge:
                # Mark all subsequent columns in the template merge range as covered
                for col in range(cell.merge.min_col + 1, cell.merge.max_col + 1):
                    covered_template_cols.add(col)
            
            # Resolve target physical column
            target_col = column_mapping.get(cell.col_index, cell.col_index)
            
            # If the cell has a merge, check if we need to shift the cell to the first visible column
            translated_merge = None
            if cell.merge:
                # Find first and last visible columns in the merge range
                first_visible_col = None
                for col in range(cell.merge.min_col, cell.merge.max_col + 1):
                    resolved = column_mapping.get(col, col)
                    if resolved is not None:
                        first_visible_col = resolved
                        break
                
                last_visible_col = None
                for col in range(cell.merge.max_col, cell.merge.min_col - 1, -1):
                    resolved = column_mapping.get(col, col)
                    if resolved is not None:
                        last_visible_col = resolved
                        break
                
                if first_visible_col is not None and last_visible_col is not None and last_visible_col >= first_visible_col:
                    target_col = first_visible_col
                    if last_visible_col > first_visible_col or cell.merge.row_span > 1:
                        translated_merge = TemplateMerge(
                            min_col=first_visible_col,
                            max_col=last_visible_col,
                            row_span=cell.merge.row_span,
                            value=cell.merge.value
                        )
            
            if target_col is None:
                # Column is hidden in this mode and does not merge into any visible column
                continue
            
            translated_cell = UnitCell(
                col_index=target_col,
                value=cell.value,
                style=cell.style,
                merge=translated_merge
            )
            translated_cells.append(translated_cell)
            
        translated_rows.append(UnitRow(
            relative_index=row.relative_index,
            height=row.height,
            cells=translated_cells
        ))
    return translated_rows
