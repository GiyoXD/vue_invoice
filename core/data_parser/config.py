import json
import os
import re
import logging
from typing import Dict, Any, Optional
from pathlib import Path

logger = logging.getLogger(__name__)


# --- File Configuration ---
INPUT_EXCEL_FILE = "JF.xlsx" # Or specific name for this format, e.g., "JF_Data_2024.xlsx"
# Specify sheet name, or None to use the active sheet
SHEET_NAME = "Invoice"
# OUTPUT_PICKLE_FILE = "invoice_data.pkl" # Example for future use

# --- Sheet Parsing Configuration ---
# Row/Column range to search for the header
# Adjusted to a more realistic range to improve performance and avoid matching stray text.
HEADER_SEARCH_ROW_RANGE = (1, 50)
HEADER_SEARCH_COL_RANGE = (1, 30) # Increased range slightly, adjust if many columns
# A pattern (string or regex) to identify a cell within the header row.
# Built dynamically from col_po keywords in TARGET_HEADERS_MAP after loading.
# PO column is the universal anchor — every table header has it.
HEADER_IDENTIFICATION_PATTERN = r"(po)"  # Fallback; rebuilt by load_and_update_mappings()


EXPECTED_HEADER_DATA_TYPES = {
    'col_po': ['string', 'numeric'],
    'col_item': ['string'], # Production Order is always a string that matches the pattern
    'col_desc': ['string'],
    'col_qty_pcs': ['numeric'],
    'col_net': ['numeric'],
    'col_gross': ['numeric'],
    'col_unit_price': ['numeric'],
    'col_amount': ['numeric'],
    'col_qty_sf': ['numeric'],
    'col_cbm': ['numeric', 'string'], # CBM can be a number or a string like '1*2*3'
    'col_dc': ['string'],
    'col_batch_no': ['string'],
    'col_line_no': ['string'],
    'col_direction': ['string'],
    'col_production_date': ['string'],
    'col_production_order_no': ['string'],
    'col_reference_code': ['string'],
    'col_level': ['string'],
    'col_pallet_count': ['numeric', 'string'],
    'col_pallet_no': ['string', 'numeric'],
    'col_pallet_id': ['string', 'numeric'],
    'col_manual_no': ['string'],
    'col_remarks': ['string'],
    'col_inv_no': ['string'],
    'col_inv_date': ['string', 'numeric', 'date'], # Invoice date can be a string or numeric date
    'col_inv_ref': ['string'],
    'col_container_no': ['string'],
    'col_unit_sf': ['numeric'],
    'col_hs_code': ['string', 'numeric'],
    'col_no': ['string', 'numeric'],
    'col_static': ['string'],
    'col_sqm': ['numeric'],
    'col_qty_header': ['string', 'numeric'],
    'col_date_recipt': ['string', 'numeric', 'date']
}

# --- Column Mapping Configuration ---
# TARGET_HEADERS_MAP is populated at import time by load_and_update_mappings()
# which reads from mapping_config.json (the single source of truth).
# Editable via the TemplateExtractor UI → "Manage Global Mappings".
TARGET_HEADERS_MAP = {}

# --- Header Validation Configuration ---
EXPECTED_HEADER_PATTERNS = {
    'col_production_order_no': [
        r'^(25|26|27)\d{5}-\d{2}$',
    ],
    'col_cbm': [
        r'^\d+(\.\d+)?\*\d+(\.\d+)?\*\d+(\.\d+)?$'
    ],
    # This pattern is now a fallback for the specific value check below
    'col_pallet_count': [
        r'^1$'
    ],
    # Pallet ID format: 2 digits + 1 letter + 8 digits (e.g. 01T26052605)
    'col_pallet_id': [
        r'^\d{2}[A-Z]\d{8}$'
    ],
    'col_remarks': [r'^\D+$'],  # Non-numeric characters only
}

EXPECTED_HEADER_VALUES = {
    # If a column header maps to 'col_pallet_count', the data value below it MUST be 1.
    # Otherwise, the column will be ignored for the 'col_pallet_count' mapping.
    'col_pallet_count': [1]
}

HEADERLESS_COLUMN_PATTERNS = {
    # If an empty header cell has data below it that looks like "number*number*number",
    # map it as the 'col_cbm' column.
    'col_cbm': [
        r'^\d+(\.\d+)?\*\d+(\.\d+)?\*\d+(\.\d+)?$',
    ],
    # TTX PO strong data pattern. Even if the header is generic (like "PO"), 
    # matching this data pattern will override it and map to col_production_order_no
    'col_production_order_no': [
        r'^(25|26|27)\d{5}-\d{2}$',
    ],
    # Pallet ID: auto-detect by data pattern (e.g. 01T26052605, 02T26052306)
    # Works even without header text, just like CBM.
    'col_pallet_id': [
        r'^\d{2}[A-Z]\d{8}$',
    ],
    'col_amount': [
        r'^\d{2,}(,\d{3})*\.\d{2}$',
    ],
}



# --- Data Extraction Configuration ---
# Choose a column likely to be empty *only* when the data rows truly end.
# 'item' is often a good candidate if item codes are always present for data rows.
STOP_EXTRACTION_ON_EMPTY_COLUMN = 'col_item'
# Safety limit for the number of data rows to read below the header within a table
MAX_DATA_ROWS_TO_SCAN = 1000

# --- Data Processing Configuration ---
# List of canonical header names for columns where values should be distributed
# CBM processing/distribution depends on the 'col_cbm' mapping above and if the column contains L*W*H strings
COLUMNS_TO_DISTRIBUTE = ["col_net", "col_gross", "col_cbm"] # Include 'col_cbm' if you want to distribute calculated CBM values

# The canonical header name of the column used for proportional distribution
DISTRIBUTION_BASIS_COLUMN = "col_qty_pcs"

# --- DAF Compounding Configuration ---
DAF_CHUNK_SIZE = 2  # How many items per group (e.g., PO1\\PO2)
DAF_INTRA_CHUNK_SEPARATOR = "/"  # Separator within a group (e.g., DOUBLE BACKSLASH)
DAF_INTER_CHUNK_SEPARATOR = "\n"  # Separator between groups (e.g., newline)

# --- Aggregation Strategy Configuration ---
# List or Tuple of *workbook filename* prefixes (case-sensitive) that trigger CUSTOM aggregation.
# Custom aggregation sums 'col_qty_sf' and 'col_amount' based ONLY on 'col_po' and 'col_item'.
# Standard aggregation sums 'col_qty_sf' based on 'col_po', 'col_item', and 'col_unit_price'.
# Example: If INPUT_EXCEL_FILE is "JF_Report_Q1.xlsx", it will match "JF".
CUSTOM_AGGREGATION_WORKBOOK_PREFIXES = () # Renamed Variable

# --- Dynamic Loading from JSON Config ---

def load_and_update_mappings():
    """
    Loads header mappings from the JSON config file and updates the
    TARGET_HEADERS_MAP dictionary. This makes the configuration dynamic.

    Reads from two sections of mapping_config.json:
    - 'header_text_mappings': explicit text → col_id overrides (e.g. template headers)
    """
    try:
        TARGET_HEADERS_MAP.clear()
        from core.database.db_manager import get_global_mapping_config
        data = get_global_mapping_config()

        # --- Source 1: header_text_mappings (explicit text → col_id) ---
        mappings = data.get('header_text_mappings', {}).get('mappings', {})
        for header, canonical in mappings.items():
            target_canonical = canonical
            if target_canonical in TARGET_HEADERS_MAP:
                if header not in TARGET_HEADERS_MAP[target_canonical]:
                    TARGET_HEADERS_MAP[target_canonical].append(header)
            else:
                TARGET_HEADERS_MAP[target_canonical] = [header]

        # --- Source 2: shipping_header_map (col_id → {keywords}) ---
        # Reverse the keywords into TARGET_HEADERS_MAP entries.
        col_defs = data.get('shipping_header_map', {})
        for col_id, props in col_defs.items():
            if not isinstance(props, dict):
                continue
            for keyword in props.get('keywords', []):
                if col_id in TARGET_HEADERS_MAP:
                    if keyword not in TARGET_HEADERS_MAP[col_id]:
                        TARGET_HEADERS_MAP[col_id].append(keyword)
                else:
                    TARGET_HEADERS_MAP[col_id] = [keyword]

        # --- Rebuild HEADER_IDENTIFICATION_PATTERN from core anchor keywords ---
        # We combine keywords from multiple essential columns to guarantee we catch any header row.
        global HEADER_IDENTIFICATION_PATTERN
        anchor_cols = ['col_po', 'col_net', 'col_gross', 'col_pallet_count', 'col_qty_pcs', 'col_qty_sf', 'col_item']
        all_keywords = []
        
        for col in anchor_cols:
            all_keywords.extend(TARGET_HEADERS_MAP.get(col, []))
            
        # Remove duplicates, escape regex special chars, and build pattern
        unique_keywords = list(set(all_keywords))
        # Filter out empty strings just in case
        unique_keywords = [kw for kw in unique_keywords if kw.strip()]
        escaped = [r"\b" + re.escape(kw) + r"\b" for kw in unique_keywords]
        
        # If we found keywords, build the pattern; otherwise fallback
        if escaped:
            HEADER_IDENTIFICATION_PATTERN = r"(" + "|".join(escaped) + r")"
        
    except json.JSONDecodeError:
        print("Warning: Could not decode mapping_config.json. Check for syntax errors. Using default TARGET_HEADERS_MAP.")
    except Exception as e:
        print(f"An unexpected error occurred while loading mapping_config.json: {e}")

# Immediately call the function to update the map when this module is imported
load_and_update_mappings()


def resolve_customer_decimals(input_filename: Optional[str]) -> Dict[str, int]:
    """
    Attempts to resolve customer blueprint from database using customer prefix
    extracted from filename (e.g. 'JF' from 'JF26024.xlsx').

    Extracts '_meta.decimals' if present, or extracts from 'styling_bundle' using
    'extract_decimal_places' from core.invoice_generator.utils.formula_formatter.

    Fallback default: {"col_amount": 2, "col_qty_sf": 2, "col_unit_price": 2}.
    """
    fallback: Dict[str, int] = {"col_amount": 2, "col_qty_sf": 2, "col_unit_price": 2}
    if not input_filename:
        logger.debug("No input_filename provided for decimal resolution. Using fallback defaults.")
        return fallback

    try:
        stem = Path(input_filename).stem

        # Check for variant locale suffixes
        locale = "KH"
        for suffix in ["_KH", "_VN"]:
            if suffix in stem.upper():
                locale = suffix.lstrip('_')
                break

        clean_stem = re.sub(r'(_KH|_VN)$', '', stem, flags=re.IGNORECASE)
        match = re.match(r'^([a-zA-Z\-_]+)', clean_stem)
        prefix = match.group(1).rstrip('-_') if match else None

        if not prefix:
            logger.debug(f"Could not extract customer prefix from filename '{input_filename}'. Using fallback decimals.")
            return fallback

        from core.database.db_manager import SessionLocal
        from core.database.repositories import BlueprintRepository
        from core.invoice_generator.utils.formula_formatter import extract_decimal_places

        db = SessionLocal()
        config_data = None
        try:
            repo = BlueprintRepository(db)
            row = repo.get_blueprint(prefix, locale)
            if not row and locale != "KH":
                row = repo.get_blueprint(prefix, "KH")
            if not row:
                variants = repo.get_customer_variants(prefix)
                if variants:
                    row = variants[0]
            if row and row.config_json:
                config_data = row.config_json
        finally:
            db.close()

        if not config_data:
            logger.debug(f"No blueprint config found in database for customer prefix '{prefix}'. Using fallback decimals.")
            return fallback

        decimals_result = dict(fallback)

        # 1. Check _meta.decimals
        meta_decimals = config_data.get("_meta", {}).get("decimals")
        if isinstance(meta_decimals, dict) and meta_decimals:
            for k, v in meta_decimals.items():
                if isinstance(v, int):
                    decimals_result[k] = v
                elif isinstance(v, (float, str)) and str(v).isdigit():
                    decimals_result[k] = int(v)
            logger.debug(f"Resolved customer decimals from _meta.decimals for '{prefix}': {decimals_result}")
            return decimals_result

        # 2. Check styling_bundle
        styling_bundle = config_data.get("styling_bundle", {})
        extracted_any = False
        if isinstance(styling_bundle, dict):
            for sheet_name, sheet_cfg in styling_bundle.items():
                if sheet_name.startswith("_") or sheet_name == "defaults":
                    continue
                if isinstance(sheet_cfg, dict):
                    columns = sheet_cfg.get("columns", {})
                    if isinstance(columns, dict):
                        for col_id, col_props in columns.items():
                            if isinstance(col_props, dict):
                                fmt = col_props.get("format")
                                if fmt:
                                    dec = extract_decimal_places(fmt)
                                    if dec is not None:
                                        decimals_result[col_id] = dec
                                        extracted_any = True

        if extracted_any:
            logger.debug(f"Resolved customer decimals from styling_bundle for '{prefix}': {decimals_result}")
        else:
            logger.debug(f"No decimal formats found in styling_bundle for '{prefix}'. Using fallback defaults.")

        return decimals_result

    except Exception as e:
        logger.warning(f"Error resolving customer decimals for '{input_filename}': {e}")
        return fallback