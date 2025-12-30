# This module contains utilities for text manipulation, date parsing, and replacement operations.

import logging
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell
from typing import List, Dict, Optional, Any
import re
import datetime

logger = logging.getLogger(__name__)

# The python-dateutil library is required for advanced date parsing.
# Install it using: pip install python-dateutil
from dateutil.parser import parse, ParserError


def excel_number_to_datetime(excel_num: Any) -> Optional[datetime.datetime]:
    """Converts an Excel date number to a Python datetime object."""
    try:
        excel_num = float(excel_num)
        # Excel's 1900 leap year bug needs to be accounted for.
        if excel_num > 59:
            excel_num -= 1
        delta = datetime.timedelta(days=excel_num - 1)
        return datetime.datetime(1900, 1, 1) + delta
    except (ValueError, TypeError):
        return None

def format_cell_as_date_smarter(cell: Cell, value: Any):
    """
    Intelligently parses a value (string, number, or datetime) into a
    datetime object and formats the cell accordingly.
    """
    parsed_date = None

    if isinstance(value, (datetime.datetime, datetime.date)):
        parsed_date = value
    elif isinstance(value, str):
        if not value.strip():
            pass
        else:
            try:
                parsed_date = parse(value, dayfirst=True)
            except (ParserError, ValueError):
                pass
    elif isinstance(value, (int, float)):
        if value >= 1:
            parsed_date = excel_number_to_datetime(value)

    if parsed_date:
        cell.value = parsed_date
        cell.number_format = "dd/mm/yyyy"
    else:
        cell.value = value

def _get_nested_data(data_dict: Dict[str, Any], path: List[Any]) -> Optional[Any]:
    """Safely retrieves a value from a nested structure of dictionaries and lists."""
    current_level = data_dict
    for key in path:
        if isinstance(current_level, dict) and key in current_level:
            current_level = current_level[key]
        elif isinstance(current_level, list):
            try:
                index = int(key)
                if 0 <= index < len(current_level):
                    current_level = current_level[index]
                else: return None
            except (ValueError, TypeError): return None
        else:
            return None
    return current_level

def find_and_replace(
    workbook: openpyxl.Workbook,
    rules: List[Dict[str, Any]],
    limit_rows: int,
    limit_cols: int,
    invoice_data: Optional[Dict[str, Any]] = None
):
    """
    A two-pass engine that handles 'exact', 'substring', and formula-based replacements.
    Pass 1: Locates all placeholders and performs simple value replacements.
    Pass 2: Uses the locations found in Pass 1 to build and apply formulas.
    """
    logger.info(f"Starting find and replace on sheets (searching range up to row {limit_rows}, col {limit_cols})")
    
    placeholder_locations: Dict[str, str] = {}
    
    simple_rules = [r for r in rules if "formula_template" not in r]
    formula_rules = [r for r in rules if "formula_template" in r]

    for sheet in workbook.worksheets:
        if sheet.sheet_state != 'visible':
            continue

        # --- PASS 1: Find all placeholder locations and apply simple replacements ---
        logger.debug("PASS 1: Locating placeholders and applying simple value replacements...")
        for row in sheet.iter_rows(max_row=limit_rows, max_col=limit_cols):
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value:
                    continue

                for rule in rules:
                    if rule.get("find") == cell.value.strip():
                        placeholder_locations[rule["find"]] = cell.coordinate
                        break

                for rule in simple_rules:
                    text_to_find = rule.get("find")
                    if not text_to_find:
                        continue
                    
                    match_mode = rule.get("match_mode", "substring")
                    is_match = (match_mode == 'exact' and cell.value.strip() == text_to_find) or \
                               (match_mode == 'substring' and text_to_find in cell.value)

                    if is_match:
                        replacement_content = None
                        if "data_path" in rule:
                            if not invoice_data: continue
                            replacement_content = _get_nested_data(invoice_data, rule["data_path"])
                            
                            # Fallback Logic
                            if replacement_content is None and "fallback_path" in rule:
                                logger.warning(f"Data not found at primary path {rule['data_path']} for '{text_to_find}', using fallback.")
                                replacement_content = _get_nested_data(invoice_data, rule["fallback_path"])
                        elif "replace" in rule:
                            replacement_content = rule["replace"]

                        if replacement_content is not None:
                            logger.debug(f"Applying rule for '{text_to_find}' at {cell.coordinate}...")
                            if rule.get("is_date", False):
                                format_cell_as_date_smarter(cell, replacement_content)
                            elif match_mode == 'exact':
                                cell.value = replacement_content
                            elif match_mode == 'substring':
                                cell.value = cell.value.replace(str(text_to_find), str(replacement_content))
                        break

        # --- PASS 2: Build and apply formula-based replacements ---
        logger.debug("PASS 2: Building and applying formula replacements...")
        if not formula_rules:
            logger.debug("No formula rules to apply")
        
        for rule in formula_rules:
            formula_template = rule["formula_template"]
            target_placeholder = rule["find"]
            
            target_cell_coord = placeholder_locations.get(target_placeholder)
            if not target_cell_coord:
                logger.warning(f"Could not find cell for formula placeholder '{target_placeholder}'. Skipping")
                continue

            dependent_placeholders = re.findall(r'(\{\[\[.*?\ এমন)\\}\\}\\)', formula_template)
            
            final_formula_str = formula_template
            all_deps_found = True
            
            for dep_placeholder in dependent_placeholders:
                dep_key = dep_placeholder.strip('{}')
                dep_coord = placeholder_locations.get(dep_key)
                
                if dep_coord:
                    final_formula_str = final_formula_str.replace(dep_placeholder, dep_coord)
                else:
                    logger.error(f"Could not find location for dependency '{dep_key}' needed by formula for '{target_placeholder}'")
                    all_deps_found = False
                    break
            
            if all_deps_found:
                final_formula_str = f"={final_formula_str}"
                logger.debug(f"Placing formula '{final_formula_str}' in cell {target_cell_coord}")
                sheet[target_cell_coord].value = final_formula_str
