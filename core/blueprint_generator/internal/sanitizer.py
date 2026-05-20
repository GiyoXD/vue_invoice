"""
Template Cleaner - Cleans raw Excel files to create blank templates.

This module is responsible for:
1. Stripping data rows from populated invoices/packing lists.
2. Preserving header rows and styling.
3. Injecting system placeholders (JFINV, JFTIME, etc.) into specific cells.
"""

import re
import logging
import hashlib
import json
from typing import Dict, Any, List, Optional, Tuple
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils import get_column_letter

from .scanner import TemplateAnalysisResult, SheetAnalysis
from core.utils.loop_profiler import tick

logger = logging.getLogger(__name__)

class ExcelTemplateSanitizer:
    """Cleans (sanitizes) raw Excel files to create reusable templates."""

    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)
        # Excel default dimensions (used to filter out empty cells with default size)
        self.DEFAULT_ROW_HEIGHT = 15.0  # Excel default row height in points
        self.DEFAULT_COL_WIDTH = 8.43   # Excel default column width in characters

    def sanitize_template(self, workbook: openpyxl.Workbook, analysis: TemplateAnalysisResult) -> Tuple[openpyxl.Workbook, Dict[str, Any]]:
        """
        Clean the provided workbook based on analysis.
        
        Args:
            workbook: openpyxl Workbook object (raw file)
            analysis: TemplateAnalysisResult
            
        Returns:
            Tuple of (Cleaned Workbook, layout_metadata_dict)
        """
        self.logger.info(f"Cleaning template for {analysis.customer_code}...")
        
        layout_metadata = {}
        
        for sheet_analysis in analysis.sheets:
            if sheet_analysis.name in workbook.sheetnames:
                ws = workbook[sheet_analysis.name]
                sheet_layout = self._clean_sheet(ws, sheet_analysis)
                layout_metadata[sheet_analysis.name] = sheet_layout
        
        # === NEW OPTIMIZATION ===
        # Delete all mapped sheets from the workbook entirely!
        # Since generation is 100% JSON-driven, the bundled XLSX only needs to contain 
        # the "Unknown/Static" sheets (like Terms & Conditions). 
        # By deleting the mapped sheets, we guarantee no customer data is leaked, 
        # and we avoid all openpyxl row-shifting overhead.
        analyzed_sheet_names = {sheet.name for sheet in analysis.sheets}
        for sheet_name in list(workbook.sheetnames):
            if sheet_name in analyzed_sheet_names:
                self.logger.info(f"Removing mapped sheet '{sheet_name}' from bundled XLSX (JSON-only mode)")
                del workbook[sheet_name]
                
        return workbook, layout_metadata

    def _determine_safe_max_column(self, ws, analysis) -> int:
        table_cur_max = 0
        if analysis.columns:
            for col in analysis.columns:
                end_col = col.col_index + (col.colspan - 1)
                if end_col > table_cur_max:
                    table_cur_max = end_col
        
        header_max_col = 0
        if analysis.header_row > 1:
            for row in ws.iter_rows(min_row=1, max_row=analysis.header_row - 1, min_col=1, max_col=40):
                for cell in row:
                    if cell.value is not None and cell.column > header_max_col:
                        header_max_col = cell.column
        
        dynamic_limit = max(table_cur_max, header_max_col) + 1
        safe_max_column = min(ws.max_column, 40, dynamic_limit)
        self.logger.info(f"    Dynamic Column Scan Limit: {safe_max_column} (Table Max: {table_cur_max}, Header Max: {header_max_col})")
        return safe_max_column



    def _capture_global_layout(self, ws, safe_max_column: int, preserved_layout: dict, analysis, table_footer_row: Optional[int]):
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
            # Do NOT fall back to defaultColWidth or a hardcoded value here —
            # that would inject phantom widths for columns the template left at
            # Excel's default, corrupting the generated output.
            if matching_dim and matching_dim.width is not None:
                preserved_layout["col_widths"][letter] = matching_dim.width

    def _capture_template_header_layout(self, ws, analysis, safe_max_column: int, preserved_layout: dict, process_and_store_style):
        for merged_range in ws.merged_cells:
            if merged_range.max_row < analysis.header_row:
                 range_str = str(merged_range)
                 top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                 val = str(top_left_cell.value) if top_left_cell.value is not None else ""
                 val_clean = val.strip()
                 preserved_layout["template_header_merges"][range_str] = val_clean
                 
        for r in range(1, analysis.header_row):
            if r in ws.row_dimensions:
                h = ws.row_dimensions[r].height
                if h is not None:
                    preserved_layout["template_header_row_heights"][str(r)] = h
                    
        if analysis.header_row > 1:
            for row in ws.iter_rows(min_row=1, max_row=analysis.header_row-1, min_col=1, max_col=safe_max_column):
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

    def _capture_template_footer(self, ws, analysis, safe_max_column: int, preserved_layout: dict, process_and_store_style, table_footer_row: Optional[int]):
        if table_footer_row is None:
            self.logger.warning(
                f"    [SKIP] Sheet '{analysis.name}': table footer (TOTAL row) not found "
                f"(scanned from row {analysis.header_row + 1} to end-of-sheet). "
                f"Treating as Form/Static sheet."
            )
            return
        else:
            end_delete = table_footer_row

        start_delete = analysis.header_row
        
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
            tick("sanitizer._delete_data_rows", sub="rows_processed")
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
    def _clean_sheet(self, ws: Worksheet, analysis: SheetAnalysis) -> Dict[str, Any]:
        """Clean a single sheet: strip data rows, inject placeholders."""
        self.logger.info(f"  Cleaning sheet: {analysis.name}")
        
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
            
        safe_max_column = self._determine_safe_max_column(ws, analysis)
        table_footer_row = self._find_table_footer_row(ws, analysis.header_row + 1, analysis, safe_max_column)
        self._capture_global_layout(ws, safe_max_column, preserved_layout, analysis, table_footer_row)
        self._capture_template_header_layout(ws, analysis, safe_max_column, preserved_layout, process_and_store_style)
        self._capture_template_footer(ws, analysis, safe_max_column, preserved_layout, process_and_store_style, table_footer_row)
        
        preserved_layout["style_palette"] = local_style_palette
        return preserved_layout



    def _build_merge_map(self, ws: Worksheet) -> Dict[str, Tuple[int, int]]:
        """
        Build a coord -> (anchor_row, anchor_col) dict for all merged ranges.
        O(total cells covered by merges) — call once per sheet, not per cell.
        """
        merge_map: Dict[str, Tuple[int, int]] = {}
        for merged_range in ws.merged_cells.ranges:
            anchor = (merged_range.min_row, merged_range.min_col)
            for row in range(merged_range.min_row, merged_range.max_row + 1):
                for col in range(merged_range.min_col, merged_range.max_col + 1):
                    coord = f"{get_column_letter(col)}{row}"
                    merge_map[coord] = anchor
        return merge_map

    def _find_table_footer_row(self, ws: Worksheet, search_start_row: int, analysis: Optional[SheetAnalysis] = None, safe_max_column: int = 20) -> Optional[int]:
        """
        Find the footer row by scanning for =SUM or =SUBTOTAL formula adjacency.

        Algorithm:
            1. Scan top-down from search_start_row to max_row (capped at 500 rows).
            2. For each row, collect column indices where cell value starts with '=SUM' or '=SUBTOTAL'.
               Skip any cell not starting with '=' for speed.
            3. If 2+ adjacent (consecutive) column indices have =SUM/=SUBTOTAL, mark row as candidate.
            4. Return the LAST (highest row number) candidate found.
            5. Fallback 1: Use the top-down scanner detection from SheetAnalysis if available.
            6. Fallback 2: if no adjacency match, try strict 'TOTAL OF:' keyword detection bottom-up.

        Args:
            ws: The worksheet to scan.
            search_start_row: The row to start scanning from (header_row + 1).
            analysis: The SheetAnalysis from the scanner (optional).

        Returns:
            The 1-based row number of the footer, or None if not found.
        """
        if search_start_row > ws.max_row:
            return None
        end_scan = min(ws.max_row, search_start_row + 500)
        if end_scan < ws.max_row:
            self.logger.warning(
                f"    [_find_table_footer_row] Scan capped at {end_scan} "
                f"(sheet has {ws.max_row} rows). Table footer may be missed if sheet is unusually large."
            )
        max_col = safe_max_column
        last_candidate = None  # Highest row with 2+ adjacent formula cells
        merge_map = self._build_merge_map(ws)  # Build once — O(merged cells)

        for row in range(search_start_row, end_scan + 1):
            tick("sanitizer._find_table_footer_row", sub="rows_scanned")
            formula_cols = []  # Column indices with =SUM or =SUBTOTAL in this row

            for col in range(1, max_col + 1):
                tick("sanitizer._find_table_footer_row", sub="cells_checked")
                cell = ws.cell(row=row, column=col)
                value = self._get_cell_value(cell, merge_map)

                if not value:
                    continue
                # Fast skip: only care about formulas
                if not value.startswith("="):
                    continue
                upper_val = value.upper()
                if "=SUM(" in upper_val or "=SUBTOTAL(" in upper_val:
                    formula_cols.append(col)

            # Check adjacency: need 2+ consecutive column indices
            if len(formula_cols) >= 2 and self._has_adjacent_pair(formula_cols):
                last_candidate = row

        if last_candidate:
            self.logger.info(f"    Footer detected via formula adjacency at row {last_candidate}")
            return last_candidate

        # --- FALLBACK 1: Scanner's Top-Down Footer Detection ---
        if analysis and analysis.footer_info and analysis.footer_info.row_num:
            self.logger.info(f"    Formula adjacency scan found nothing. Falling back to scanner footer info at row {analysis.footer_info.row_num}.")
            return analysis.footer_info.row_num

        # --- FALLBACK 2: Original 'TOTAL' keyword scan (bottom-up), but strict ---
        self.logger.info("    Formula adjacency scan and scanner info found nothing. Falling back to strict TOTAL keyword scan.")
        fallback_candidate = None
        scan_limit_bottom_up = max(search_start_row, ws.max_row - 500)
        for row in range(ws.max_row, scan_limit_bottom_up - 1, -1):
            tick("sanitizer._find_table_footer_row", sub="fallback_rows_scanned")
            for col in range(1, max_col + 1):
                tick("sanitizer._find_table_footer_row", sub="fallback_cells_checked")
                cell = ws.cell(row=row, column=col)
                value = self._get_cell_value(cell, merge_map)
                if value:
                    val_upper = value.upper().strip()
                    # Use exact/near-exact match to prevent picking up random sentence "Total Net Weight"
                    if val_upper in ["TOTAL", "TOTAL:", "TOTAL OF:", "TOTAL OF", "TOTAL：", "TOTAL AMOUNT", "TOTAL AMOUNT:", "TOTAL AMOUNT："] or val_upper.startswith("TOTAL OF") or val_upper.startswith("TOTAL AMOUNT"):
                        fallback_candidate = row
                        break
            if fallback_candidate:
                break

        if fallback_candidate:
            self.logger.warning(f"    Using strict TOTAL keyword fallback at row {fallback_candidate}.")
        return fallback_candidate

    def _has_adjacent_pair(self, sorted_cols: list) -> bool:
        """
        Check if a sorted list of column indices contains at least one adjacent pair.

        Args:
            sorted_cols: List of 1-based column indices (already in order from left-to-right scan).

        Returns:
            True if any two consecutive entries differ by exactly 1.
        """
        for i in range(len(sorted_cols) - 1):
            if sorted_cols[i + 1] - sorted_cols[i] == 1:
                return True
        return False

    def _get_cell_value(self, cell, merge_map: Optional[Dict[str, Tuple[int, int]]] = None) -> Optional[str]:
        """
        Return the string value of a cell, resolving merged-cell anchors.

        Args:
            cell: The openpyxl Cell or MergedCell to read.
            merge_map: Optional pre-built coord->anchor dict from _build_merge_map.
                       When supplied, anchor lookup is O(1). Without it, falls back
                       to the original O(ranges) linear scan (safe for one-off calls).
        """
        if isinstance(cell, MergedCell) or cell.value is None:
            ws = cell.parent
            coord = cell.coordinate
            if merge_map is not None:
                # O(1) lookup via pre-built map
                anchor = merge_map.get(coord)
                if anchor:
                    top_left = ws.cell(row=anchor[0], column=anchor[1])
                    if top_left.value is not None:
                        return str(top_left.value)
                return None
            # Fallback: original O(ranges) scan for one-off callers
            for merged_range in ws.merged_cells.ranges:
                if coord in merged_range:
                    top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                    if top_left_cell.value is not None:
                        return str(top_left_cell.value)
            return None

        return str(cell.value)

    def _should_record_empty_cell(self, ws: Worksheet, row: int, col: int) -> bool:
        """
        Check if an empty cell should be recorded based on non-default dimensions.
        
        An empty cell is worth recording if:
        - Its row has a height different from Excel's default (15.0 pt)
        - Its column has a width different from Excel's default (8.43 chars)
        
        This prevents recording thousands of empty cells with only default styling.
        """

        
        # Check row height
        if row in ws.row_dimensions:
            height = ws.row_dimensions[row].height
            if height is not None and height != self.DEFAULT_ROW_HEIGHT:
                return True
        
        # Check column width
        for dim in ws.column_dimensions.values():
            if dim.min <= col <= dim.max:
                if dim.width is not None and dim.width != self.DEFAULT_COL_WIDTH:
                    return True
                break
        
        return False

    def _capture_cell_style(self, cell: Cell, is_empty: bool = False) -> Optional[Dict[str, Any]]:
        """
        Capture font, alignment, border, and fill styles from a cell.
        Returns None if strict mode is on and the cell has effectively default style.
        
        Args:
            cell: The cell to analyze.
            is_empty: If True, we ignore "setup" styles like Font Name/Size and only capture 
                      visible modifiers like Borders, Fills, Bold, etc.
        """
        style = {}
        has_significant_style = False
        
        # 1. Font
        if cell.font:
            # Filter out default font props to save space
            font_data = {}
            
            # "Setup" styles (name, size) are invisible on empty cells — skip them.
            # Bold/italic are visible modifiers kept regardless.
            if not is_empty:
                # 1. Capture font name if not default (Calibri/Arial)
                if cell.font.name and cell.font.name not in ["Calibri", "Arial"]:
                    font_data["name"] = cell.font.name

                # 2. Capture font size if it's NOT (Calibri 11)
                if cell.font.size is not None:
                    is_calibri = (cell.font.name == "Calibri")
                    is_size_11 = (cell.font.size in [11.0, 11])
                    if not (is_calibri and is_size_11):
                        font_data["size"] = cell.font.size
            
            if cell.font.bold: font_data["bold"] = True
            if cell.font.italic: font_data["italic"] = True
            if cell.font.color and hasattr(cell.font.color, "rgb"): # Capture colors
                 color_val = self._serialize_color(cell.font.color)
                 if color_val and color_val not in ("00000000", "FF000000"): # Skip black/auto
                     font_data["color"] = color_val

            if font_data:
                style["font"] = font_data
                has_significant_style = True
            
        # 2. Alignment
        if cell.alignment:
            align_data = {}
            # Only capture if NOT default (general/bottom)
            if cell.alignment.horizontal and cell.alignment.horizontal != 'general':
                align_data["horizontal"] = cell.alignment.horizontal
            if cell.alignment.vertical and cell.alignment.vertical != 'bottom': # bottom is default in Excel? usually.
                align_data["vertical"] = cell.alignment.vertical
            if cell.alignment.wrap_text:
                align_data["wrap_text"] = True
                
            if align_data:
                style["alignment"] = align_data
                has_significant_style = True
            
        # 3. Fill (Background)
        if cell.fill and cell.fill.fill_type and cell.fill.fill_type != "none":
            if hasattr(cell.fill, "start_color"):
                 color_val = self._serialize_color(cell.fill.start_color)
                 # Skip default "none" or white fills
                 if color_val and color_val not in ["00000000", "FFFFFFFF"]:
                     style["fill"] = {
                         "type": cell.fill.fill_type,
                         "color": color_val
                     }
                     has_significant_style = True
             
        # 4. Border
        if cell.border:
             border_data = {}
             # Only capture if there is an actual border style
             if cell.border.left and cell.border.left.style: border_data["left"] = cell.border.left.style
             if cell.border.right and cell.border.right.style: border_data["right"] = cell.border.right.style
             if cell.border.top and cell.border.top.style: border_data["top"] = cell.border.top.style
             if cell.border.bottom and cell.border.bottom.style: border_data["bottom"] = cell.border.bottom.style
             
             if border_data:
                 style["border"] = border_data
                 has_significant_style = True
             
        # 5. Number Format
        # Only relevant if content exists, usually? 
        # Actually user might pre-format a column for dates. 
        # But for "Pure Empty Junk" check, maybe skip? 
        # User said "modify more property like border or change width and height". 
        # Number format is invisible until data is typed.
        # Let's keep strictness: If empty, ignore number format too?
        # Safe bet: Capture it. It's rare to have Custom Num Format on junk cells.
        if cell.number_format and cell.number_format != "General":
            style["number_format"] = cell.number_format
            has_significant_style = True
        
        # If nothing significant was captured, return None to save massive JSON space
        return style if has_significant_style else None

    def _serialize_color(self, color) -> Optional[str]:
        """Try to extract RGB hex string from Color object."""
        if color is None: return None
        if hasattr(color, "rgb") and color.rgb:
            # openpyxl rgb is usually "AARRGGBB" or "RRGGBB"
            # We treat it as string
            if isinstance(color.rgb, str):
                return color.rgb
        if hasattr(color, "theme") and color.theme is not None:
            return f"theme-{color.theme}"
        return None

