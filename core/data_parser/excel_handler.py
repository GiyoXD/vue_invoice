# --- START OF FULL FILE: excel_handler.py ---

import openpyxl
import os
import logging # Using logging is better than print for info/errors

# Basic config moved to main.py, logger will inherit settings
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ExcelHandler:
    """Handles loading and accessing data from Excel files using openpyxl."""
    def __init__(self, file_path_or_buffer):
        self.is_buffer = hasattr(file_path_or_buffer, "read")
        
        if not self.is_buffer:
            if not os.path.exists(file_path_or_buffer):
                logging.error(f"File not found: {file_path_or_buffer}")
                raise FileNotFoundError(f"The file '{file_path_or_buffer}' was not found.")
                
        self.file_path = file_path_or_buffer
        self.workbook = None
        self.sheet = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def load_sheet(self, sheet_name=None, data_only=True, fallback_to_active=True):
        """
        Loads the workbook and a specific sheet.

        Args:
            sheet_name (str, optional): Name of the sheet. Defaults to None (active sheet).
            data_only (bool, optional): Get cell values (True) or formulas (False). Defaults to True.
            fallback_to_active (bool, optional): If sheet_name not found, fallback to active sheet. Defaults to True.

        Returns:
            openpyxl.worksheet.worksheet.Worksheet: The loaded sheet object.
        """
        try:
            self.workbook = openpyxl.load_workbook(self.file_path, data_only=data_only)

            if sheet_name:
                if sheet_name in self.workbook.sheetnames:
                    self.sheet = self.workbook[sheet_name]
                else:
                    normalized_target = sheet_name.strip().lower()
                    sheet_map = {name.strip().lower(): name for name in self.workbook.sheetnames}
                    if normalized_target in sheet_map:
                        self.sheet = self.workbook[sheet_map[normalized_target]]
                    elif fallback_to_active:
                        logging.warning(
                            f"Sheet '{sheet_name}' not found in workbook. "
                            f"Falling back to active sheet '{self.workbook.active.title}'."
                        )
                        self.sheet = self.workbook.active
                    else:
                        raise KeyError(f"Sheet '{sheet_name}' not found. Available sheets: {self.workbook.sheetnames}")
            else:
                self.sheet = self.workbook.active

            logging.debug(f"Loaded sheet '{self.sheet.title}' ({self.sheet.max_row} rows, {self.sheet.max_column} cols)")
            
            # Add diagnostic info for MOTO files
            if "MOTO" in str(self.file_path):
                logging.warning(f"MOTO FILE DETECTED - Large max_row might cause performance issues: {self.sheet.max_row}")
            
            return self.sheet
        except Exception as e:
            target_name = getattr(self.file_path, 'name', 'in-memory buffer') if self.is_buffer else str(self.file_path)
            logging.error(f"Failed to load workbook/sheet from '{target_name}': {e}", exc_info=True)
            self.close()
            raise

    def get_sheet(self):
        """Returns the currently loaded sheet object."""
        if not self.sheet:
            logging.warning("Sheet not loaded. Call load_sheet() first.")
        return self.sheet

    def close(self):
        """Closes the workbook if it's open."""
        # openpyxl doesn't require explicit closing for reading,
        # but closing the workbook object might release resources sooner.
        if self.workbook:
            try:
                # Although load_workbook doesn't keep the file open,
                # calling close releases the workbook object from memory sooner.
                self.workbook.close()
                logging.info(f"Closed workbook object reference for: {self.file_path}")
            except Exception as e:
                # This shouldn't typically happen on read-only workbooks
                logging.warning(f"Exception while closing workbook reference (this is usually okay): {e}")
            finally:
                 self.workbook = None
                 self.sheet = None


# --- END OF FULL FILE: excel_handler.py ---