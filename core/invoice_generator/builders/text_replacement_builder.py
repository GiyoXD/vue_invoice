import logging
import openpyxl
from typing import Dict, Any

# Import the core replacement engine
from ..utils.text import find_and_replace

logger = logging.getLogger(__name__)

class TextReplacementBuilder:
    """
    A builder class responsible for handling all text replacement tasks within the invoice.
    """
    def __init__(self, workbook: openpyxl.Workbook, invoice_data: Dict[str, Any]):
        self.workbook = workbook
        self.invoice_data = invoice_data

    def build(self):
        """
        Executes all configured text replacement tasks.
        """
        self._replace_placeholders()
        self._run_daf_specific_replacement()

    def _replace_placeholders(self):
        """Defines and runs the data-driven placeholder replacement task."""
        logger.info("Running placeholder replacement task (within A1:N14)")
        header_rules = [
            {"find": "JFINV", "data_path": ["processed_tables_data", "1", "col_inv_no", 0], "match_mode": "exact"},
            {"find": "JFTIME", "data_path": ["processed_tables_data", "1", "col_inv_date", 0], "is_date": True, "match_mode": "exact"},
            {"find": "JFREF", "data_path": ["processed_tables_data", "1", "col_inv_ref", 0], "match_mode": "exact"},
            {"find": "[[CUSTOMER_NAME]]", "data_path": ["customer_info", "name"], "match_mode": "exact"},
            {"find": "[[CUSTOMER_ADDRESS]]", "data_path": ["customer_info", "address"], "match_mode": "exact"}
        ]
        find_and_replace(
            workbook=self.workbook,
            rules=header_rules,
            limit_rows=14,
            limit_cols=14,
            invoice_data=self.invoice_data
        )
        logger.info("Finished placeholder replacement task")

    def _run_daf_specific_replacement(self):
        """Defines and runs the hardcoded, DAF-specific replacement task."""
        logger.info("Running DAF-specific replacement task (within 50x16 grid)")
        daf_rules = [
            {"find": "BINH PHUOC", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BAVET, SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BAVET,SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BAVET, SVAYRIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BINH DUONG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "FCA  BAVET,SVAYRIENG", "replace": "DAF BAVET", "match_mode": "exact"},
            {"find": "FCA: BAVET,SVAYRIENG", "replace": "DAF: BAVET", "match_mode": "exact"},
            {"find": "DAF  BAVET,SVAYRIENG", "replace": "DAF BAVET", "match_mode": "exact"},
            {"find": "DAF: BAVET,SVAYRIENG", "replace": "DAF: BAVET", "match_mode": "exact"},
            {"find": "SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "PORT KLANG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "HCM", "replace": "BAVET", "match_mode": "exact"},
            {"find": "DAP", "replace": "DAF", "match_mode": "substring"},
            {"find": "FCA", "replace": "DAF", "match_mode": "substring"},
            {"find": "CIF", "replace": "DAF", "match_mode": "substring"},
        ]
        find_and_replace(
            workbook=self.workbook,
            rules=daf_rules,
            limit_rows=200,
            limit_cols=16
        )
        logger.info("Finished DAF-specific replacement task")
