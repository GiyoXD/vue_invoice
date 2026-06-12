"""
Template Cleaner - Cleans raw Excel files to create blank templates.

This module is responsible for:
1. Stripping data rows from populated invoices/packing lists.
2. Preserving header rows and styling.
"""

import logging
from typing import List
import openpyxl
from openpyxl.cell.cell import MergedCell

logger = logging.getLogger(__name__)

class ExcelTemplateSanitizer:
    """Cleans (sanitizes) raw Excel files to create reusable templates."""

    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)

    def sanitize_template(self, workbook: openpyxl.Workbook, analyzed_sheet_names: List[str]) -> openpyxl.Workbook:
        """
        Clean the provided workbook by removing/clearing analyzed sheets.
        
        Args:
            workbook: openpyxl Workbook object (raw file)
            analyzed_sheet_names: List of sheet names that were analyzed
            
        Returns:
            Cleaned Workbook object
        """
        self.logger.info("Cleaning template sheets...")
        
        # === NEW OPTIMIZATION ===
        # Delete all mapped sheets from the workbook entirely!
        # Since generation is 100% JSON-driven, the bundled XLSX only needs to contain 
        # the "Unknown/Static" sheets (like Terms & Conditions). 
        # By deleting the mapped sheets, we guarantee no customer data is leaked, 
        # and we avoid all openpyxl row-shifting overhead.
        analyzed_set = set(analyzed_sheet_names)
        remaining_sheets = list(workbook.sheetnames)
        for sheet_name in list(workbook.sheetnames):
            if sheet_name in analyzed_set:
                if len(remaining_sheets) > 1:
                    self.logger.info(f"Removing mapped sheet '{sheet_name}' from bundled XLSX (JSON-only mode)")
                    del workbook[sheet_name]
                    remaining_sheets.remove(sheet_name)
                else:
                    self.logger.info(f"Keeping last sheet '{sheet_name}' in bundled XLSX but clearing cell values to prevent data leak")
                    ws = workbook[sheet_name]
                    for row in ws.iter_rows():
                        for cell in row:
                            if not isinstance(cell, MergedCell):
                                cell.value = None
                
        return workbook

