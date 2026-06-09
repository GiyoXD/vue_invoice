import re
import logging
import hashlib
import json
from typing import Dict, List, Any, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils import get_column_letter

from core.utils.loop_profiler import tick
from core.blueprint_generator.utils.footer_scanner import FooterInfo
from .models import ZoneBoundaries, ColumnInfo

logger = logging.getLogger(__name__)

class TemplateScanner:
    """Scans static template zones (Zone 1 & 3): content outside the table area."""

    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)
        # Excel default dimensions (used to filter out empty cells with default size)
        self.DEFAULT_ROW_HEIGHT = 15.0  # Excel default row height in points
        self.DEFAULT_COL_WIDTH = 8.43   # Excel default column width in characters

    def scan_static_content(self, worksheet: Worksheet, boundaries: ZoneBoundaries, footer_info: Optional[FooterInfo], columns: List[ColumnInfo], sheet_name: str) -> Dict[str, Any]:
        """
        Scan static content zones outside the table area (Zone 1 & 3).
        Returns the layout metadata dictionary.
        """
        self.logger.info(f"  Scanning static template layout: {sheet_name}")
        
        preserved_layout = {
            "template_header_merges": {},
            "template_header_row_heights": {},
            "template_header_content": {},
            "template_header_styles": {},
            "template_footer_rows": [],
            "style_palette": {},
            "col_widths": {},
            "template_header_images": [],
            "template_footer_images": []
        }
        
        local_style_palette = {}
        
        def process_and_store_style(style_dict: Dict[str, Any]) -> str:
            style_str = json.dumps(style_dict, sort_keys=True)
            style_hash = "style_" + hashlib.md5(style_str.encode('utf-8')).hexdigest()[:8]
            if style_hash not in local_style_palette:
                local_style_palette[style_hash] = style_dict
            return style_hash
            
        safe_max_column = self._determine_safe_max_column(worksheet, columns, boundaries)
        table_footer_row = boundaries.footer_row if boundaries.footer_row is not None else (footer_info.row_num if footer_info else None)
        
        self._capture_global_layout(worksheet, safe_max_column, preserved_layout)
        self._capture_template_header_layout(worksheet, boundaries, safe_max_column, preserved_layout, process_and_store_style)
        self._capture_template_footer(worksheet, boundaries, safe_max_column, preserved_layout, process_and_store_style, table_footer_row, sheet_name)
        
        preserved_layout["style_palette"] = local_style_palette
        return preserved_layout

    def _determine_safe_max_column(self, ws: Worksheet, columns: List[ColumnInfo], boundaries: ZoneBoundaries) -> int:
        table_cur_max = 0
        if columns:
            for col in columns:
                end_col = col.col_index + (col.colspan - 1)
                if end_col > table_cur_max:
                    table_cur_max = end_col
        
        header_max_col = 0
        if boundaries.header_row > 1:
            for row in ws.iter_rows(min_row=1, max_row=boundaries.header_row - 1, min_col=1, max_col=40):
                for cell in row:
                    if cell.value is not None and cell.column > header_max_col:
                        header_max_col = cell.column
        
        dynamic_limit = max(table_cur_max, header_max_col) + 1
        safe_max_column = min(ws.max_column, 40, dynamic_limit)
        self.logger.info(f"    Dynamic Column Scan Limit: {safe_max_column} (Table Max: {table_cur_max}, Header Max: {header_max_col})")
        return safe_max_column

    def _capture_global_layout(self, ws: Worksheet, safe_max_column: int, preserved_layout: dict):
        # Cache grouped dimension ranges (openpyxl stores <col min="1" max="5" width="20"/>
        # under a single dict key). We must iterate them to find the matching range for each column.
        dim_ranges = list(ws.column_dimensions.values())

        for c in range(1, safe_max_column + 1):
            letter = get_column_letter(c)

            # Find the dimension object whose range covers this column index
            matching_dim = None
            for dim in dim_ranges:
                if dim.min <= c <= dim.max:
                    matching_dim = dim
                    break

            # Only record widths that are EXPLICITLY set in the worksheet.
            if matching_dim and matching_dim.width is not None:
                preserved_layout["col_widths"][letter] = matching_dim.width

    def _capture_template_header_layout(self, ws: Worksheet, boundaries: ZoneBoundaries, safe_max_column: int, preserved_layout: dict, process_and_store_style):
        for merged_range in ws.merged_cells:
            if merged_range.max_row < boundaries.header_row:
                 range_str = str(merged_range)
                 top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                 val = str(top_left_cell.value) if top_left_cell.value is not None else ""
                 val_clean = val.strip()
                 preserved_layout["template_header_merges"][range_str] = val_clean
                 
        for r in range(1, boundaries.header_row):
            if r in ws.row_dimensions:
                h = ws.row_dimensions[r].height
                if h is not None:
                    preserved_layout["template_header_row_heights"][str(r)] = h
                    
        if boundaries.header_row > 1:
            for row in ws.iter_rows(min_row=1, max_row=boundaries.header_row-1, min_col=1, max_col=safe_max_column):
                for cell in row:
                    coord = cell.coordinate
                    is_empty = (cell.value is None)
                    
                    if not is_empty:
                         val_str = str(cell.value)
                         if val_str.startswith('='):
                             val_str = re.sub(r'\[\d+\]', '', val_str)
                         preserved_layout["template_header_content"][coord] = val_str
                    
                    style_data = self._capture_cell_style(cell, is_empty=is_empty)
                    
                    if is_empty and not style_data and not self._should_record_empty_cell(ws, cell.row, cell.column):
                        continue
                        
                    if style_data:
                        style_id = process_and_store_style(style_data)
                        if style_id not in preserved_layout["template_header_styles"]:
                            preserved_layout["template_header_styles"][style_id] = []
                        preserved_layout["template_header_styles"][style_id].append(coord)

    def _capture_template_footer(self, ws: Worksheet, boundaries: ZoneBoundaries, safe_max_column: int, preserved_layout: dict, process_and_store_style, table_footer_row: Optional[int], sheet_name: str):
        if table_footer_row is None:
            self.logger.warning(
                f"    [SKIP] Sheet '{sheet_name}': table footer (TOTAL row) not found "
                f"(scanned from row {boundaries.header_row + 1} to end-of-sheet). "
                f"Treating as Form/Static sheet."
            )
            return
        else:
            end_delete = table_footer_row

        start_delete = boundaries.header_row
        
        preserved_layout["template_header_images"] = []
        preserved_layout["template_footer_images"] = []
        
        self.logger.info(f"    Capturing footer data (Rows {end_delete + 1} to EOF)")
        
        template_footer_merges = []
        footer_merge_map_by_row = {}
        # Capture merges for JSON metadata
        for merged_range in list(ws.merged_cells):
            m_min_row, m_min_col, m_max_row, m_max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
            
            if m_min_row > start_delete:
                merge_tuple = (m_min_row, m_min_col, m_max_row, m_max_col)
                if m_min_row not in footer_merge_map_by_row:
                    footer_merge_map_by_row[m_min_row] = []
                footer_merge_map_by_row[m_min_row].append(merge_tuple)

        template_footer_heights = []
        template_footer_rows = []
        # CAP THE FOOTER SCAN to prevent a corrupted max_row (e.g. 1,000,000) from hanging the loop.
        # An invoice footer is rarely more than 100 rows.
        current_max_row = min(ws.max_row, end_delete + 100)
        
        for r in range(end_delete + 1, current_max_row + 1):
            tick("scanner._capture_template_footer", sub="rows_processed")
            rel_r = r - (end_delete + 1)
            row_dict = {
                "relative_index": rel_r,
                "height": None,
                "merges": [],
                "cells": []
            }
            
            if r in ws.row_dimensions:
                h = ws.row_dimensions[r].height
                if h is not None:
                    row_dict["height"] = h
                    template_footer_heights.append((r, h))
                    
            if r in footer_merge_map_by_row:
                for (old_min_r, min_c, old_max_r, max_c) in footer_merge_map_by_row[r]:
                    top_left_cell = ws.cell(row=r, column=min_c)
                    val = str(top_left_cell.value) if top_left_cell.value is not None else ""
                    val_clean = val.strip()
                    
                    row_dict["merges"].append({
                        "min_col": min_c,
                        "max_col": max_c,
                        "row_span": old_max_r - old_min_r + 1,
                        "value": val_clean
                    })
                    
            has_content_or_style = False
            for c in range(1, safe_max_column + 1):
                cell = ws.cell(row=r, column=c)
                is_empty = (cell.value is None)
                
                style_data = self._capture_cell_style(cell, is_empty=is_empty)
                
                if is_empty and not style_data and not self._should_record_empty_cell(ws, r, c):
                    continue
                    
                cell_dict = {"col_index": c}
                has_content_or_style = True
                
                if not is_empty:
                    val_str = str(cell.value)
                    if val_str.startswith('='):
                        val_str = re.sub(r'\[\d+\]', '', val_str)
                    cell_dict["value"] = val_str
                    
                if style_data:
                    style_id = process_and_store_style(style_data)
                    cell_dict["style_id"] = style_id
                    
                row_dict["cells"].append(cell_dict)
                
            if row_dict["height"] is not None or row_dict["merges"] or has_content_or_style:
                template_footer_rows.append(row_dict)
                
        preserved_layout["template_footer_rows"] = template_footer_rows



    def _should_record_empty_cell(self, ws: Worksheet, row: int, col: int) -> bool:
        if row in ws.row_dimensions:
            height = ws.row_dimensions[row].height
            if height is not None and height != self.DEFAULT_ROW_HEIGHT:
                return True
        for dim in ws.column_dimensions.values():
            if dim.min <= col <= dim.max:
                if dim.width is not None and dim.width != self.DEFAULT_COL_WIDTH:
                    return True
                break
        return False

    def _capture_cell_style(self, cell: Cell, is_empty: bool = False) -> Optional[Dict[str, Any]]:
        style = {}
        has_significant_style = False
        
        # 1. Font
        if cell.font:
            font_data = {}
            if not is_empty:
                if cell.font.name and cell.font.name not in ["Calibri", "Arial"]:
                    font_data["name"] = cell.font.name
                if cell.font.size is not None:
                    is_calibri = (cell.font.name == "Calibri")
                    is_size_11 = (cell.font.size in [11.0, 11])
                    if not (is_calibri and is_size_11):
                        font_data["size"] = cell.font.size
            if cell.font.bold: font_data["bold"] = True
            if cell.font.italic: font_data["italic"] = True
            if cell.font.color and hasattr(cell.font.color, "rgb"):
                 color_val = self._serialize_color(cell.font.color)
                 if color_val and color_val not in ("00000000", "FF000000"):
                     font_data["color"] = color_val

            if font_data:
                style["font"] = font_data
                has_significant_style = True
            
        # 2. Alignment
        if cell.alignment:
            align_data = {}
            if cell.alignment.horizontal and cell.alignment.horizontal != 'general':
                align_data["horizontal"] = cell.alignment.horizontal
            if cell.alignment.vertical and cell.alignment.vertical != 'bottom':
                align_data["vertical"] = cell.alignment.vertical
            if cell.alignment.wrap_text:
                align_data["wrap_text"] = True
            if align_data:
                style["alignment"] = align_data
                has_significant_style = True
            
        # 3. Fill
        if cell.fill and cell.fill.fill_type and cell.fill.fill_type != "none":
            if hasattr(cell.fill, "start_color"):
                 color_val = self._serialize_color(cell.fill.start_color)
                 if color_val and color_val not in ["00000000", "FFFFFFFF"]:
                     style["fill"] = {
                         "type": cell.fill.fill_type,
                         "color": color_val
                     }
                     has_significant_style = True
             
        # 4. Border
        if cell.border:
             border_data = {}
             if cell.border.left and cell.border.left.style: border_data["left"] = cell.border.left.style
             if cell.border.right and cell.border.right.style: border_data["right"] = cell.border.right.style
             if cell.border.top and cell.border.top.style: border_data["top"] = cell.border.top.style
             if cell.border.bottom and cell.border.bottom.style: border_data["bottom"] = cell.border.bottom.style
             if border_data:
                 style["border"] = border_data
                 has_significant_style = True
             
        # 5. Number Format
        if cell.number_format and cell.number_format != "General":
            style["number_format"] = cell.number_format
            has_significant_style = True
        
        return style if has_significant_style else None

    def _serialize_color(self, color) -> Optional[str]:
        if color is None: return None
        if hasattr(color, "rgb") and color.rgb:
            if isinstance(color.rgb, str):
                return color.rgb
        if hasattr(color, "theme") and color.theme is not None:
            return f"theme-{color.theme}"
        return None
