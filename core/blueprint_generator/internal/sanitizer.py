"""
Template Cleaner - Cleans raw Excel files to create blank templates.

This module is responsible for:
1. Stripping data rows from populated invoices/packing lists.
2. Preserving header rows and styling.
3. Injecting system placeholders (JFINV, JFTIME, etc.) into specific cells.
"""

import logging
from typing import Dict, Any, Tuple
import openpyxl
from openpyxl.cell.cell import MergedCell

from .scanner import TemplateAnalysisResult

logger = logging.getLogger(__name__)

class ExcelTemplateSanitizer:
    """Cleans (sanitizes) raw Excel files to create reusable templates."""

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
                # Retrieve the static layout that was captured by the TemplateScanner during the scanning phase
                if sheet_analysis.static_layout:
                    layout_metadata[sheet_analysis.name] = sheet_analysis.static_layout
                else:
                    self.logger.warning(f"  Missing static layout for sheet: {sheet_analysis.name}")
        
        # === NEW OPTIMIZATION ===
        # Delete all mapped sheets from the workbook entirely!
        # Since generation is 100% JSON-driven, the bundled XLSX only needs to contain 
        # the "Unknown/Static" sheets (like Terms & Conditions). 
        # By deleting the mapped sheets, we guarantee no customer data is leaked, 
        # and we avoid all openpyxl row-shifting overhead.
        analyzed_sheet_names = {sheet.name for sheet in analysis.sheets}
        remaining_sheets = list(workbook.sheetnames)
        for sheet_name in list(workbook.sheetnames):
            if sheet_name in analyzed_sheet_names:
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
                
        return workbook, layout_metadata
