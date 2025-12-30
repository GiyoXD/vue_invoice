import logging
from typing import Dict, Any, List
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from .base_processor import SheetProcessor

logger = logging.getLogger(__name__)

from copy import copy

class PlaceholderProcessor(SheetProcessor):
    """
    A lightweight processor that performs simple text replacement in an Excel sheet.
    It does NOT perform complex table generation, row insertion, or style application.
    It strictly replaces {{PLACEHOLDERS}} with values from the input data.
    """

    def __init__(
        self,
        template_workbook: Workbook,
        output_workbook: Workbook,
        template_worksheet: Worksheet,
        output_worksheet: Worksheet,
        sheet_name: str,
        sheet_config: Dict[str, Any],
        config_loader: Any,
        data_source_indicator: str,
        invoice_data: Dict[str, Any],
        cli_args: Any = None,
        final_grand_total_pallets: int = 0
    ):
        # We accept the same signature as other processors for compatibility,
        # even if we don't use all arguments.
        self.template_workbook = template_workbook
        self.output_workbook = output_workbook
        self.template_worksheet = template_worksheet
        self.output_worksheet = output_worksheet
        self.sheet_name = sheet_name
        self.sheet_config = sheet_config
        self.invoice_data = invoice_data
        
        # For metadata logging
        self.replacements_log = []
        self.header_info = {}

    def process(self) -> bool:
        """
        Executes the placeholder replacement logic.
        Returns True if successful, False otherwise.
        """
        try:
            logger.info(f"Starting Placeholder Processing for sheet: {self.sheet_name}")
            
            # 0. Copy content from template to output
            # Since WorkbookBuilder creates a blank sheet, we must copy the template content first.
            for row in self.template_worksheet.iter_rows():
                for cell in row:
                    new_cell = self.output_worksheet.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = copy(cell.number_format)
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)
            
            # Copy column dimensions
            for col_dim in self.template_worksheet.column_dimensions.values():
                self.output_worksheet.column_dimensions[col_dim.index] = copy(col_dim)
                
            # Copy row dimensions
            for row_dim in self.template_worksheet.row_dimensions.values():
                self.output_worksheet.row_dimensions[row_dim.index] = copy(row_dim)
                
            # Copy merges
            for merge_range in self.template_worksheet.merged_cells.ranges:
                self.output_worksheet.merge_cells(str(merge_range))
            
            # 1. Flatten the invoice data for easier lookup
            # We assume the data might be nested, but for placeholders we usually want flat keys.
            # Strategy: Use the entire invoice_data as the context.
            # If the user wants specific sub-dictionaries, they can reference them in the placeholder?
            # For simplicity, let's flatten the top level and 'metadata' if it exists.
            
            context = self.invoice_data.copy()
            if 'metadata' in context and isinstance(context['metadata'], dict):
                context.update(context['metadata'])
            
            # Also support the "data_map" from config if it exists, to map specific keys
            # But the user requested "simple replacement", so maybe just direct key lookup is best.
            # Let's stick to direct key lookup from the JSON data first.
            
            # 2. Iterate through all cells in the output worksheet
            # Since we are modifying the output worksheet (which is a copy of the template),
            # we just need to find and replace.
            
            replacements_count = 0
            
            for row in self.output_worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        original_value = cell.value
                        new_value = original_value
                        
                        # Check for placeholders pattern: {{KEY}}
                        # We can do a simple iteration over keys if the data is small,
                        # or use regex if we want to be fancy. 
                        # Given the requirement for simplicity, let's iterate over the context keys
                        # ONLY if the cell looks like it has a placeholder.
                        
                        if "{{" in new_value and "}}" in new_value:
                            for key, val in context.items():
                                # We only replace scalar values (str, int, float)
                                if isinstance(val, (str, int, float)):
                                    placeholder = f"{{{{{key}}}}}" # e.g. {{INVOICE_NUM}}
                                    if placeholder in new_value:
                                        new_value = new_value.replace(placeholder, str(val))
                                        
                            if new_value != original_value:
                                cell.value = new_value
                                replacements_count += 1
                                self.replacements_log.append({
                                    "sheet": self.sheet_name,
                                    "cell": cell.coordinate,
                                    "original": original_value,
                                    "new": new_value
                                })

            logger.info(f"Completed Placeholder Processing. Replaced {replacements_count} values.")
            return True

        except Exception as e:
            logger.error(f"Error in PlaceholderProcessor for sheet {self.sheet_name}: {e}")
            return False
