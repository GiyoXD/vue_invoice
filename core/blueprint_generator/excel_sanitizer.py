"""
Template Cleaner - Cleans raw Excel files to create blank templates.

This module is responsible for:
1. Stripping data rows from populated invoices/packing lists.
2. Preserving header rows and styling.
3. Injecting system placeholders (JFINV, JFTIME, etc.) into specific cells.
"""

import logging
from typing import Dict, Any, List, Optional, Tuple
import openpyxl
import base64
from io import BytesIO
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, MergedCell

from .excel_scanner import TemplateAnalysisResult, SheetAnalysis

logger = logging.getLogger(__name__)

class ExcelTemplateSanitizer:
    """Cleans (sanitizes) raw Excel files to create reusable templates."""
    
    # Text to Placeholder mappings
    # If a cell contains valid data that looks like metadata, we replace it with these placeholders
    PLACEHOLDERS = {
        "invoice no": "JFINV",
        "inv no": "JFINV",
        "date": "JFTIME",
        "date:": "JFTIME",
        "ref": "JFREF",
        "reference": "JFREF",
        "invoice ref": "JFREF"
    }

    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)

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
        
        # === FIX FOR CORRUPTION ===
        # Clear all images from sheets before returning.
        # Images in openpyxl can cause corruption when the workbook is saved
        # after row deletions. We've already captured image data to metadata.
        for sheet in workbook.worksheets:
            if hasattr(sheet, '_images'):
                sheet._images = []
            # Also clear drawings which can cause issues
            if hasattr(sheet, '_charts'):
                sheet._charts = []
                
        return workbook, layout_metadata

    def _clean_sheet(self, ws: Worksheet, analysis: SheetAnalysis) -> Dict[str, Any]:
        """Clean a single sheet: strip data rows, inject placeholders."""
        self.logger.info(f"  Cleaning sheet: {analysis.name}")
        
        preserved_layout = {
            "header_merges": [],
            "header_row_heights": {},
            "header_content": {},
            "header_styles": {},
            "footer_merges": [],
            "footer_row_heights": {},
            "footer_content": {},
            "footer_styles": {},
            "col_widths": {},
            "header_images": [],
            "footer_images": []
        }
        
        # --- CAPTURE GLOBAL LAYOUT (Column Widths) ---
        from openpyxl.utils import get_column_letter
        for c in range(1, ws.max_column + 1):
            letter = get_column_letter(c)
            if letter in ws.column_dimensions:
                w = ws.column_dimensions[letter].width
                # width can be None for default
                if w is not None:
                     preserved_layout["col_widths"][letter] = w
        
        # --- CAPTURE HEADER LAYOUT & CONTENT ---
        # Capture strictly ABOVE the header row (Metadata area)
        
        # 1. Header Merges
        from openpyxl.utils import get_column_letter
        for merged_range in ws.merged_cells:
            if merged_range.max_row < analysis.header_row:
                 range_str = str(merged_range)
                 preserved_layout["header_merges"].append(range_str)
                 
        # 2. Header Row Heights
        for r in range(1, analysis.header_row):
            if r in ws.row_dimensions:
                h = ws.row_dimensions[r].height
                if h is not None:
                    preserved_layout["header_row_heights"][str(r)] = h
                    
        # 3. Header Content & Styles
        for r in range(1, analysis.header_row):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                coord = cell.coordinate
                
                # Content
                if cell.value is not None:
                    preserved_layout["header_content"][coord] = str(cell.value)
                    
                # Style (capture for all cells in range, to get borders/fills even if empty)
                # But to save space, maybe only if it has 'interesting' properties? 
                # For now, capture all to ensure fidelity.
                preserved_layout["header_styles"][coord] = self._capture_cell_style(cell)
        
        # 1. Find Footer (Total Row) FIRST, before messing with indices?
        # Actually indices are stable if we haven't deleted yet.
        # Scan bottom-up for "Total" + "=SUM"
        # We search from header_row + 1
        footer_start_row = self._find_footer_start(ws, analysis.header_row + 1)
        
        if not footer_start_row:
             self.logger.warning(f"    Footer not detected in {analysis.name}. Assuming this is a Form/Static sheet (no dynamic table body).")
             self.logger.warning("    Skipping row deletion to preserve content.")
             # Set end_delete to header_row - 1 so rows_to_delete becomes 0
             end_delete = analysis.header_row - 1
             footer_start_row = ws.max_row # Treat end of sheet as footer start for image capture purposes
        else:
             end_delete = footer_start_row

        # 2. Delete ENTIRE Table (Header to Footer inclusive)
        # User requested: "remove from the header to the footer entirely"
        # This removes Header, Data Body, and the Footer row found.
        # Remaining content (like Signature) will shift up.
        
        start_delete = analysis.header_row
        
        if end_delete < start_delete:
            rows_to_delete = 0
        else:
            rows_to_delete = end_delete - start_delete + 1
        
        
        # 3. SKIP Image Capture - was causing corruption
        # Images are cleared in sanitize_template() before saving
        preserved_layout["header_images"] = []
        preserved_layout["footer_images"] = []
        
        if rows_to_delete > 0:
            self.logger.info(f"    Deleting ENTIRE TABLE: {rows_to_delete} rows (Rows {start_delete}-{end_delete} | Header {start_delete} to Footer {end_delete})")
            
            # --- PRE-DELETION MERGE HANDLING ---
            # 1. Ranges IN the deletion zone: Destroy them.
            # 2. Ranges BELOW the deletion zone (Footer): Store & Destroy, then Restore after deletion (to prevent corruption).
            
            footer_merges_to_restore = []
            
            # Iterate over a copy of the list because we modify it
            for merged_range in list(ws.merged_cells):
                m_min_row, m_min_col, m_max_row, m_max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
                
                # Case A: Merge intersects deletion zone -> Destroy (Ghost Merge Prevention)
                if (m_min_row <= end_delete and m_max_row >= start_delete):
                    self.logger.info(f"    Unmerging intersecting range {merged_range} before deletion.")
                    ws.unmerge_cells(str(merged_range))
                    
                # Case B: Merge is strictly BELOW deletion zone (Footer) -> Store & Unmerge
                elif m_min_row > end_delete:
                    self.logger.info(f"    Storing & Temporarily Unmerging footer range {merged_range} to preserve format.")
                    footer_merges_to_restore.append((m_min_row, m_min_col, m_max_row, m_max_col))
                    ws.unmerge_cells(str(merged_range))

            # Case C: Capture Row Heights & Content for Footer (strictly below deletion zone)
            footer_heights_to_restore = []
            footer_content_dist = {} # Stores { "A25": "Value" } using SHIFTED coordinates
            footer_styles_dist = {} # Stores { "A25": {style_dict} }
            
            # Use max_row from before deletion
            current_max_row = ws.max_row
            from openpyxl.utils import get_column_letter
            
            for r in range(end_delete + 1, current_max_row + 1):
                # 1. Capture Height
                if r in ws.row_dimensions:
                    h = ws.row_dimensions[r].height
                    if h is not None:
                        footer_heights_to_restore.append((r, h))
                        
                # 2. Capture Content & Styles
                shifted_r = r - rows_to_delete
                if shifted_r < 1: continue 
                
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(row=r, column=c)
                    new_coord = f"{get_column_letter(c)}{shifted_r}"
                    
                    if cell.value is not None:
                         footer_content_dist[new_coord] = str(cell.value)
                    
                    # Capture Style mapped to NEW coordinate
                    footer_styles_dist[new_coord] = self._capture_cell_style(cell)
                        
            # Save captured content to metadata immediately (calculated for post-shift)
            preserved_layout["footer_content"] = footer_content_dist
            preserved_layout["footer_styles"] = footer_styles_dist

            # --- DELETION ---
            ws.delete_rows(start_delete, amount=rows_to_delete)
            
            # --- RESTORE FOOTER MERGES ---
            # Shift rows up by rows_to_delete
            for (old_min_r, min_c, old_max_r, max_c) in footer_merges_to_restore:
                new_min_r = old_min_r - rows_to_delete
                new_max_r = old_max_r - rows_to_delete
                
                # Sanity check
                if new_min_r < 1: 
                    continue
                    
                self.logger.info(f"    Restoring footer merge at rows {new_min_r}-{new_max_r} (was {old_min_r}-{old_max_r})")
                ws.merge_cells(start_row=new_min_r, start_column=min_c, end_row=new_max_r, end_column=max_c)
                
                # Save to metadata (using openpyxl range string format, e.g. "A40:G40")
                # We need to construct the range string manually or use helper
                from openpyxl.utils import get_column_letter
                range_str = f"{get_column_letter(min_c)}{new_min_r}:{get_column_letter(max_c)}{new_max_r}"
                preserved_layout["footer_merges"].append(range_str)

            # --- RESTORE FOOTER ROW HEIGHTS ---
            for (old_r, height) in footer_heights_to_restore:
                new_r = old_r - rows_to_delete
                if new_r < 1: continue
                
                # self.logger.info(f"    Restoring footer row height {height} at row {new_r} (was {old_r})")
                ws.row_dimensions[new_r].height = height
                preserved_layout["footer_row_heights"][str(new_r)] = height

        else:
            self.logger.warning("    Row calculation result <= 0? Check header/footer detection.")
            
        # 4. Inject Placeholders in Metadata Area (Above Header)
        # Scan cells above header_row for metadata labels
        self._inject_placeholders(ws, analysis.header_row)
        
        return preserved_layout

    def _inject_placeholders(self, ws: Worksheet, header_row: int):
        """Scan area above header for metadata labels and inject placeholders next to them."""
        # Limit scan to reasonable header area
        scan_rows = header_row - 1
        if scan_rows < 1: return
        
        # Track cells we've already injected into to prevent duplicates
        injected_cells = set()
        
        for row in range(1, scan_rows + 1):
            for col in range(1, 15): # Scan first 15 columns
                cell = ws.cell(row=row, column=col)
                value = self._get_cell_value(cell)
                
                if value:
                    # Skip if this cell already has a JF placeholder
                    if value.upper().startswith("JF"):
                        continue
                    
                    val_lower = value.lower().strip()
                    # Keep original for some checks, clean version for others
                    val_clean = val_lower.rstrip(':').rstrip('.').strip()
                    # Also normalize spaces and dots for matching
                    val_normalized = val_clean.replace('.', ' ').replace('  ', ' ')

                    # Fuzzy / Key matching
                    placeholder = None
                    
                    # INVOICE NO detection
                    if "invoice no" in val_normalized or "inv no" in val_normalized or val_normalized == "inv":
                        placeholder = "JFINV"
                    
                    # DATE detection (but avoid long strings like "Shipping Date")
                    elif "date" in val_clean and len(val_clean) < 15:
                        placeholder = "JFTIME"
                    
                    # REF NO detection - RESTRICTIVE matching
                    # Only match SHORT labels that look like "Ref No:", "Ref.", etc.
                    # Max length 15 chars to avoid matching addresses or long text
                    elif len(val_clean) < 15 and (
                        any(pattern in val_normalized for pattern in ["ref no", "ref num"]) or  # "Ref No", "Ref. No"
                        val_normalized in ["ref", "reference"] or  # Exact match "Ref" or "Reference"
                        (val_clean.endswith("ref") and len(val_clean) < 10) or  # "Cust Ref", "Inv Ref"
                        ("ref" in val_clean and ("inv" in val_clean or "cust" in val_clean))  # "Inv Ref No"
                    ):
                        placeholder = "JFREF"
                    
                    if placeholder:
                        # Look for value in next cell (right)
                        target_cell = None
                        
                        # Try immediate right
                        c1 = ws.cell(row=row, column=col + 1)
                        if not isinstance(c1, MergedCell):
                            target_cell = c1
                        
                        if target_cell:
                            # Skip if we've already injected here
                            if target_cell.coordinate in injected_cells:
                                continue
                            
                            # Skip if target already has a JF placeholder
                            target_value = self._get_cell_value(target_cell)
                            if target_value and target_value.upper().startswith("JF"):
                                continue
                            
                            self.logger.info(f"    Found metadata label '{value}' at {cell.coordinate}. Injecting {placeholder} at {target_cell.coordinate}")
                            target_cell.value = placeholder
                            injected_cells.add(target_cell.coordinate)

    def _find_footer_start(self, ws: Worksheet, search_start_row: int) -> Optional[int]:
        """
        Find the start of the footer by looking broadly for the 'Total' row.
        Algorithm:
        - Iterate backwards from max_row.
        - STRICT MATCH: 'total' + '=SUM' formula.
        - RELAXED MATCH (Fallback): 'total' + numeric/currency content.
        """
        # Scan backwards
        best_candidate = None
        
        for row in range(ws.max_row, search_start_row, -1):
            has_total_keyword = False
            has_sum_formula = False
            has_value = False
            
            for col in range(1, min(20, ws.max_column + 1)): # Scan first 20 cols
                cell = ws.cell(row=row, column=col)
                value = self._get_cell_value(cell)
                
                if value:
                    val_lower = value.lower()
                    if "total" in val_lower:
                        has_total_keyword = True
                        
                    if val_lower.startswith("=sum"):
                        has_sum_formula = True
                    
                    # Check for simple numeric/currency values often found in total rows
                    if any(c in val_lower for c in ['$', '€', '£']) or value.replace('.','',1).isdigit():
                        has_value = True
                        
            # STRICT MATCH
            if has_total_keyword and has_sum_formula:
                return row
            
            # Record Candidate for Fallback (First one from bottom up is likely the grand total)
            if has_total_keyword and best_candidate is None:
                 best_candidate = row
                 
        # If strict match failed, return best candidate (just "Total" keyword)
        if best_candidate:
            self.logger.warning(f"    Strict footer detection failed (Total + Formula). Using relaxed match at row {best_candidate}.")
            return best_candidate
            
        return None

    def _get_cell_value(self, cell) -> Optional[str]:
        if isinstance(cell, MergedCell) or cell.value is None:
            return None
        return str(cell.value)

    def _capture_cell_style(self, cell: Cell) -> Dict[str, Any]:
        """Capture font, alignment, border, and fill styles from a cell."""
        style = {}
        
        # 1. Font
        if cell.font:
            style["font"] = {
                "name": cell.font.name,
                "size": cell.font.size,
                "bold": cell.font.bold,
                "italic": cell.font.italic,
                "color": self._serialize_color(cell.font.color)
            }
            
        # 2. Alignment
        if cell.alignment:
            style["alignment"] = {
                "horizontal": cell.alignment.horizontal,
                "vertical": cell.alignment.vertical,
                "wrap_text": cell.alignment.wrap_text
            }
            
        # 3. Fill (Background)
        if cell.fill and hasattr(cell.fill, "start_color"):
             style["fill"] = {
                 "type": cell.fill.fill_type,
                 "color": self._serialize_color(cell.fill.start_color)
             }
             
        # 4. Border (Simplified - only capturing style presence for now)
        if cell.border:
             style["border"] = {
                 "left": cell.border.left.style if cell.border.left else None,
                 "right": cell.border.right.style if cell.border.right else None,
                 "top": cell.border.top.style if cell.border.top else None,
                 "bottom": cell.border.bottom.style if cell.border.bottom else None
             }
             
        # 5. Number Format
        style["number_format"] = cell.number_format
        
        return style

    def _serialize_color(self, color) -> Optional[str]:
        """Try to extract RGB hex string from Color object."""
        if not color: return None
        if hasattr(color, "rgb") and color.rgb:
            # openpyxl rgb is usually "AARRGGBB" or "RRGGBB"
            # We treat it as string
            if isinstance(color.rgb, str):
                return color.rgb
        return None

    def _capture_images_from_sheet(self, ws: Worksheet, header_limit: int, footer_start: int, delete_count: int) -> Tuple[List[Dict], List[Dict]]:
        """
        Capture images from the sheet, classifying them as Header or Footer.
        Returns (header_images, footer_images).
        Footer images have their row coordinates SHIFTED up by delete_count.
        """
        header_imgs = []
        footer_imgs = []
        
        # Access internal images list (standard way in openpyxl for now)
        if not hasattr(ws, "_images"):
            return header_imgs, footer_imgs
            
        for img in ws._images:
            # Determine anchor row (0-indexed in openpyxl, convert to 1-based for comparison)
            # anchor type varies (OneCellAnchor, TwoCellAnchor)
            # But usually has _from
            try:
                row_idx = img.anchor._from.row
                col_idx = img.anchor._from.col
                # col_off = img.anchor._from.colOff
                # row_off = img.anchor._from.rowOff
                
                row_num = row_idx + 1 # 1-based
                
                img_data = self._serialize_image(img)
                if not img_data: continue
                
                # Add anchor info
                from openpyxl.utils import get_column_letter
                anchor_col = get_column_letter(col_idx + 1)
                
                # Header Image
                if row_num < header_limit:
                    img_data["anchor"] = f"{anchor_col}{row_num}"
                    # Allow fine definition? For now just cell anchor.
                    # Ideally we want offsets too.
                    img_data["anchor_details"] = {
                        "col": col_idx, "row": row_idx,
                        "colOff": getattr(img.anchor._from, "colOff", 0),
                        "rowOff": getattr(img.anchor._from, "rowOff", 0)
                    }
                    header_imgs.append(img_data)
                    
                # Footer Image (Strictly below deletion zone)
                # If deletion goes up to footer_start, then Footer starts at footer_start + 1 ??
                # In my logic: end_delete = footer_start. So Footer is > footer_start.
                elif row_num > footer_start:
                    # Shift row
                    new_row_num = row_num - delete_count
                    if new_row_num < 1: continue
                    
                    img_data["anchor"] = f"{anchor_col}{new_row_num}"
                    img_data["anchor_details"] = {
                        "col": col_idx, "row": row_idx - delete_count, # Shifted internal row
                        "colOff": getattr(img.anchor._from, "colOff", 0),
                        "rowOff": getattr(img.anchor._from, "rowOff", 0)
                    }
                    footer_imgs.append(img_data)
                    
            except Exception as e:
                self.logger.warning(f"Failed to capture image: {e}")
                
        return header_imgs, footer_imgs

    def _serialize_image(self, img) -> Optional[Dict]:
        """Convert image to base64 dict."""
        try:
            # Read binary data
            # img.ref is a file-like object or bytes?
            # openpyxl Image has ._data? or we read from path?
            # Actually img._data might cover it if loaded.
            data = None
            if hasattr(img, "_data"): # internal blob
                 data = img._data()
            elif hasattr(img, "ref") and hasattr(img.ref, "read"):
                 data = img.ref.read()
            elif hasattr(img, "path"): # path to file
                 with open(img.path, "rb") as f:
                     data = f.read()
                     
            if not data: return None
            
            b64_str = base64.b64encode(data).decode('utf-8')
            
            return {
                "width": img.width,
                "height": img.height,
                "format": img.format,
                "data": b64_str
            }
        except (ValueError, OSError, AttributeError, Exception) as e:
            # Catch 'I/O operation on closed file' (ValueError) and others
            self.logger.warning(f"Error serializing image: {e}")
            return None
