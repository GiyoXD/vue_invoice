import logging
import copy
from typing import List, Dict, Any, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

# Reuse text replacement logic
from ..utils.text import find_and_replace

logger = logging.getLogger(__name__)

class JsonTemplateStateBuilder:
    """
    A builder responsible for reconstructing the template state from a JSON configuration
    rather than scanning a live Excel worksheet.
    
    This replaces the scanning logic in TemplateStateBuilder with JSON parsing logic.
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
        
        # State structures (same as TemplateStateBuilder)
        self.header_state: List[List[Dict[str, Any]]] = []
        self.footer_state: List[List[Dict[str, Any]]] = []
        self.header_merged_cells: List[str] = []
        self.footer_merged_cells: List[str] = []
        self.row_heights: Dict[int, float] = {}
        self.column_widths: Dict[int, float] = {}
        
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
        
        # Log of text replacements
        self.replacements_log: List[Dict[str, str]] = []

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
        """Converting the coordinate-based JSON into row-based state lists."""
        logger.info("Parsing JSON layout data for template reconstruction...")
        
        # 1. Parse Dimensions and Basic Props
        # Note: JSON 'col_widths' keys are letters 'A', 'B'...
        col_widths = self.layout_data.get('col_widths', {})
        for col_letter, width in col_widths.items():
            self.column_widths[column_index_from_string(col_letter)] = width
            
        # 2. Parse Header State
        # 'header_content' is {"A1": {...}, "B1": {...}}
        header_content = self.layout_data.get('header_content', {})
        header_styles = self.layout_data.get('header_styles', {})
        self.header_merged_cells = self.layout_data.get('header_merges', [])
        
        self.header_state, self.header_end_row = self._build_state_grid(header_content, header_styles, is_header=True)
        
        # Load header row heights
        header_row_heights = self.layout_data.get('header_row_heights', {})
        for r_str, h in header_row_heights.items():
            self.row_heights[int(r_str)] = h

        # 3. Parse Footer State
        footer_content = self.layout_data.get('footer_content', {})
        footer_styles = self.layout_data.get('footer_styles', {})
        self.footer_merged_cells = self.layout_data.get('footer_merges', [])
        
        self.footer_state, self.template_footer_end_row = self._build_state_grid(footer_content, footer_styles, is_header=False)
        
        # Determine footer start row
        # In JSON, footer keys are absolute (e.g., "A50"). We need to find the Min row in footer keys.
        if footer_content or footer_styles:
            all_keys = list(footer_content.keys()) + list(footer_styles.keys())
            min_r = float('inf')
            for k in all_keys:
                try:
                    _, r = coordinate_from_string(k)
                    if r < min_r: min_r = r
                except: pass
            self.template_footer_start_row = min_r if min_r != float('inf') else -1
        
        # Load footer row heights
        footer_row_heights = self.layout_data.get('footer_row_heights', {})
        for r_str, h in footer_row_heights.items():
            self.row_heights[int(r_str)] = h
            
        # Update max_col
        if self.column_widths:
            self.max_col = max(self.column_widths.keys())
            
        logger.info(f"JSON Parse Complete: Header Ends Row {self.header_end_row}, Footer Starts Row {self.template_footer_start_row}")

    def _build_state_grid(self, content_map: Dict, style_map: Dict, is_header: bool) -> tuple:
        """
        Convert coordinate maps (A1: val) to list-of-lists grid.
        Returns (state_grid, max_row_index)
        """
        if not content_map and not style_map:
            return [], 0

        # Find bounds
        all_coords = set(content_map.keys()) | set(style_map.keys())
        if not all_coords:
            return [], 0
            
        rows = set()
        cols = set()
        for coord in all_coords:
            try:
                c, r = coordinate_from_string(coord)
                rows.add(r)
                cols.add(column_index_from_string(c))
            except:
                continue
                
        if not rows: return [], 0
        
        min_r = min(rows)
        max_r = max(rows)
        max_c = max(cols) if cols else 1
        
        # Ensure we cover at least the columns defined in widths or arbitrary max
        final_max_c = max(max_c, max(self.column_widths.keys()) if self.column_widths else 0)
        
        grid = []
        
        # Iterate row by row
        for r in range(min_r, max_r + 1):
            row_data = []
            for c in range(1, final_max_c + 1):
                col_letter = get_column_letter(c)
                coord = f"{col_letter}{r}"
                
                # Extract value
                raw_val = content_map.get(coord, None)
                
                # Extract style dict
                style_dict = style_map.get(coord, {})
                
                # Convert style dict to OpenPyXL objects
                cell_info = {
                    'value': raw_val,
                    'font': self._create_font(style_dict.get('font')),
                    'fill': self._create_fill(style_dict.get('fill')),
                    'border': self._create_border(style_dict.get('border')),
                    'alignment': self._create_alignment(style_dict.get('alignment')),
                    'number_format': style_dict.get('number_format', 'General')
                }
                row_data.append(cell_info)
            grid.append(row_data)
            
        return grid, max_r

    # --- Style Factory Methods ---
    def _create_font(self, d: Dict) -> Optional[Font]:
        if not d: return None
        # Handle color dict/str
        color = d.get('color')
        # If color is dict (RGB), extract rgb
        if isinstance(color, dict) and 'rgb' in color:
             color = color['rgb']
        elif isinstance(color, dict) and 'theme' in color:
             # Simplify theme colors to None or black for now unless we look up theme
             color = None
             
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
        if not d.get('type'): return None
        # Simplification: mostly dealing with solid fills usually
        # The serializer saves 'color' as '00000000' usually for transparent
        # We need check how sanitizer saves it. 
        # For now, instantiate basic PatternFill
        fgColor = d.get('color')
        if fgColor == '00000000': fgColor = None # Transparent
        
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
    
    def restore_header_only(self, target_worksheet: Worksheet, actual_num_cols: int = None):
        """Restores ONLY the header to a new clean worksheet."""
        if self.debug:
            logger.debug(f"Restoring header from JSON state")
        
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
                self._write_cell(target_cell, cell_info)
                
        # Restore header merges
        for merge_str in self.header_merged_cells:
            self._apply_merge(target_worksheet, merge_str)
            
        # Restore dimensions
        for r_idx in range(self.min_row, self.header_end_row + 1):
             if r_idx in self.row_heights:
                 target_worksheet.row_dimensions[r_idx].height = self.row_heights[r_idx]
                 
        for c_idx, w in self.column_widths.items():
            target_worksheet.column_dimensions[get_column_letter(c_idx)].width = w

    def restore_footer_only(self, target_worksheet: Worksheet, footer_start_row: int, actual_num_cols: int = None):
        """Restores ONLY the footer to the new worksheet."""
        if self.debug:
            logger.debug(f"Restoring footer from JSON state at row {footer_start_row}")
            
        offset = footer_start_row - self.template_footer_start_row
        
        for row_idx, row_data in enumerate(self.footer_state):
            actual_row = self.template_footer_start_row + row_idx + offset
            
            for col_idx, cell_info in enumerate(row_data):
                template_col = col_idx + self.min_col
                output_col = self._get_mapped_column(template_col)
                
                if output_col is None: continue
                
                target_cell = target_worksheet.cell(row=actual_row, column=output_col)
                self._write_cell(target_cell, cell_info)

        # Restore footer merges
        for merge_str in self.footer_merged_cells:
             self._apply_merge(target_worksheet, merge_str, offset=offset)
             
        # Restore footer heights
        for r_idx in range(self.template_footer_start_row, self.template_footer_end_row + 1):
            if r_idx in self.row_heights:
                target_worksheet.row_dimensions[r_idx + offset].height = self.row_heights[r_idx]

    def _write_cell(self, cell, info):
        if info['value'] is not None:
            cell.value = info['value']
        if info['font']: cell.font = copy.copy(info['font'])
        if info['fill']: cell.fill = copy.copy(info['fill'])
        if info['border']: cell.border = copy.copy(info['border'])
        if info['alignment']: cell.alignment = copy.copy(info['alignment'])
        if info['number_format']: cell.number_format = info['number_format']

    def _apply_merge(self, ws, merge_str, offset=0):
        try:
            from openpyxl.utils.cell import range_boundaries
            min_col, min_row, max_col, max_row = range_boundaries(merge_str)
            
            # Apply offset
            min_row += offset
            max_row += offset
            
            # Apply column mapping (simplified logic - assume continuous blocks for now or strict mapping)
            # For robust mapping, we need the full logic from TemplateStateBuilder, but for this step
            # we will trust the caller handles the complex "removed column" cases or add that later.
            # Just straightforward mapping here:
            
            mapped_min_col = self._get_mapped_column(min_col)
            mapped_max_col = self._get_mapped_column(max_col)
            
            if mapped_min_col and mapped_max_col:
                new_range = f"{get_column_letter(mapped_min_col)}{min_row}:{get_column_letter(mapped_max_col)}{max_row}"
                ws.merge_cells(new_range)
        except Exception as e:
            logger.warning(f"Failed to apply merge {merge_str}: {e}")

    def apply_text_replacements(self, replacement_rules: list, invoice_data: dict = None) -> int:
        """Mirror of TemplateStateBuilder.apply_text_replacements"""
        # We can implement the exact same logic here since we have the same structure
        changes = 0
        
        def process_grid(grid):
            c = 0
            for row in grid:
                for cell in row:
                    val = cell.get('value')
                    if val and isinstance(val, str):
                        # Simple exact/placeholder check for now
                        # Full rule engine logic:
                        new_val, fmt_change, term = self._apply_rules_to_cell(val, replacement_rules, invoice_data)
                        if new_val != val:
                            cell['value'] = new_val
                            if fmt_change: cell['number_format'] = 'General'
                            c += 1
            return c

        changes += process_grid(self.header_state)
        changes += process_grid(self.footer_state)
        return changes

    def _apply_rules_to_cell(self, text: str, rules: list, invoice_data: dict = None) -> tuple:
        # Simplified version of logic from TemplateStateBuilder
        if not text: return text, False, None
        
        for rule in rules:
            find_text = rule.get('find', '')
            if not find_text: continue
            
            # Resolve replacement
            replacement = None
            if 'data_path' in rule and invoice_data:
                # Need _resolve_data_path logic
                replacement = self._resolve_path(invoice_data, rule['data_path'])
            elif 'replace' in rule:
                replacement = rule['replace']
            
            if replacement is None: continue
            
            if text.strip() == find_text:
                return str(replacement), True, find_text
            elif find_text in text:
                 return text.replace(find_text, str(replacement)), True, find_text
                 
        return text, False, None

    def _resolve_path(self, data, path):
        curr = data
        try:
            for k in path:
                if isinstance(curr, list): curr = curr[int(k)]
                else: curr = curr.get(k)
            return curr
        except:
            return None
