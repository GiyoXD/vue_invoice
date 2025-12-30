# invoice_generator/utils/text_replace_utils.py
"""
Utility functions for text replacement tasks in invoices.
These are wrapper functions that use the core find_and_replace engine.
"""

import openpyxl
from typing import Dict, Any
from .text import find_and_replace
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def run_invoice_header_replacement_task(workbook: openpyxl.Workbook, invoice_data: Dict[str, Any]):
    """Defines and runs the data-driven header replacement task."""
    logger.info("\n--- Running Invoice Header Replacement Task (within A1:N14) ---")
    header_rules = [
        {"find": "JFINV", "data_path": ["invoice_info", "inv_no"], "fallback_path": ["processed_tables_data", "1", "col_inv_no", 0], "match_mode": "exact"},
        # This rule will now correctly handle any date format coming from your data
        {"find": "JFTIME", "data_path": ["invoice_info", "inv_date"], "fallback_path": ["processed_tables_data", "1", "col_inv_date", 0], "is_date": True, "match_mode": "exact"},
        {"find": "JFREF", "data_path": ["invoice_info", "inv_ref"], "fallback_path": ["processed_tables_data", "1", "col_inv_ref", 0], "match_mode": "exact"},
        {"find": "[[CUSTOMER_NAME]]", "data_path": ["customer_info", "name"], "match_mode": "exact"},
        {"find": "[[CUSTOMER_ADDRESS]]", "data_path": ["customer_info", "address"], "match_mode": "exact"}
    ]
    find_and_replace(
        workbook=workbook,
        rules=header_rules,
        limit_rows=14,
        limit_cols=14,
        invoice_data=invoice_data
    )
    logger.info("--- Finished Invoice Header Replacement Task ---")


def run_DAF_specific_replacement_task(workbook: openpyxl.Workbook):
    """Defines and runs the hardcoded, DAF-specific replacement task."""
    logger.info("\n--- Running DAF-Specific Replacement Task (within 50x16 grid) ---")
    DAF_rules = [
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
        workbook=workbook,
        rules=DAF_rules,
        limit_rows=200,
        limit_cols=16
    )
    logger.info("--- Finished DAF-Specific Replacement Task ---")
