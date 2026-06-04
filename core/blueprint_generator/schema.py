"""
Blueprint Schema - Single Source of Truth for Invoice Generation Business Logic.

This module defines the business rules for:
1. Column Identification (keywords -> column definition)
2. Sheet Classification (name -> data source)
3. Column Formatting (id -> excel number format)

NOTE: Additional column definitions are loaded dynamically from
`mapping_config.json` (the `shipping_header_map` section) at class
load time. Add new system columns there instead of hardcoding here.
"""

import json
import logging
import re
from dataclasses import dataclass, field
from typing import List, Dict, Set, Optional

from core.utils.loop_profiler import tick

logger = logging.getLogger(__name__)

@dataclass
class ColumnDefinition:
    """Defines how a specific column type behaves."""
    id: str                  # Internal System ID (e.g., 'col_qty_pcs')
    keywords: List[str]      # Header keywords for matching (e.g., ['pcs', 'quantity'])
    excel_format: str = "@"  # Default Excel number format (Text)
    width: float = 15.0      # Standard width fallback
    priority: int = 10       # Higher priority matches first (if needed)

class BlueprintSchema:
    """Central registry of blueprint schema and definitions."""

    # 1. Sheet Classification Rules (Loaded dynamically)
    AGGREGATION_SHEETS: Set[str] = set()
    PROCESSED_TABLES_SHEETS: Set[str] = set()
    ALLOWED_SEARCH_SHEETS: Set[str] = set()

    # Global column scan limit to prevent infinite loops on malformed sheets with ghost columns (e.g. 16384)
    MAX_SCAN_COLUMN: int = 50
    
    # Footer Scanning Bounds (How far above/below the TOTAL row to search for the HS Code)
    FOOTER_HS_SEARCH_WINDOW: int = 50

    # 2. Column Definitions
    # These replace the hardcoded HEADER_MAPPINGS in excel_scanner.py
    # and the hardcoded format checks.
    _BASE_COLUMNS: Dict[str, ColumnDefinition] = {
        "col_static": ColumnDefinition(id="col_static", keywords=[], width=24.71),
        "col_po": ColumnDefinition(id="col_po", keywords=[], width=28.0),
        "col_item": ColumnDefinition(id="col_item", keywords=[], width=22.14),
        "col_desc": ColumnDefinition(id="col_desc", keywords=[], width=26.0),
        "col_qty_header": ColumnDefinition(id="col_qty_header", keywords=[]),
        "col_qty_pcs": ColumnDefinition(id="col_qty_pcs", keywords=[], excel_format="#,##0"),
        "col_qty_sf": ColumnDefinition(id="col_qty_sf", keywords=[], excel_format="#,##0.00"),
        "col_unit_sf": ColumnDefinition(id="col_unit_sf", keywords=[], excel_format="#,##0.00"),
        "col_unit_price": ColumnDefinition(id="col_unit_price", keywords=[], excel_format="#,##0.00"),
        "col_amount": ColumnDefinition(id="col_amount", keywords=[], excel_format="#,##0.00", width=18.0),
        "col_net": ColumnDefinition(id="col_net", keywords=[], excel_format="#,##0.00"),
        "col_gross": ColumnDefinition(id="col_gross", keywords=[], excel_format="#,##0.00"),
        "col_cbm": ColumnDefinition(id="col_cbm", keywords=[], excel_format="0.00"),
        "col_no": ColumnDefinition(id="col_no", keywords=[]),
        "col_pallet_count": ColumnDefinition(id="col_pallet_count", keywords=[]),
        "col_pallet_id": ColumnDefinition(id="col_pallet_id", keywords=[]),
        "col_dc": ColumnDefinition(id="col_dc", keywords=[], width=12.0),
        "col_container_no": ColumnDefinition(id="col_container_no", keywords=[], width=18.0),
        "col_remarks": ColumnDefinition(id="col_remarks", keywords=[], width=20.0),
        "col_sqm": ColumnDefinition(id="col_sqm", keywords=[], excel_format="#,##0.00", width=16.0),
        "col_hs_code": ColumnDefinition(id="col_hs_code", keywords=[]),
    }

    # Initialize COLUMNS to be dynamically populated/modified by load_dynamic_columns()
    COLUMNS: Dict[str, ColumnDefinition] = {}

    # Pre-built keyword index: {keyword_lower: ColumnDefinition}
    # Built once by _rebuild_keyword_index(), called after _load_from_config()
    _KEYWORD_INDEX: Dict[str, 'ColumnDefinition'] = {}

    @classmethod
    def get_column_by_keyword(cls, header_text: str) -> Optional[ColumnDefinition]:
        """
        Identify a column definition based on header text.
        Returns the Best Match (or None).
        
        Uses pre-built _KEYWORD_INDEX for O(1) lookup instead of
        linear scan through all COLUMNS × keywords.
        """
        if not header_text:
            return None
            
        header_lower = header_text.lower().strip()
        
        # 1. O(1) exact match via pre-built index
        result = cls._KEYWORD_INDEX.get(header_lower)
        if result:
            tick("schema.get_column_by_keyword", sub="index_hits")
            return result
        
        tick("schema.get_column_by_keyword", sub="index_misses")
        
        # 2. Smart fallback for HS Code (regex-like: both 'hs' and 'code' present)
        header_clean = re.sub(r'[^a-z0-9]', '', header_lower)
        if 'hs' in header_clean and 'code' in header_clean:
            return cls.COLUMNS.get("col_hs_code")
                    
        return None

    @classmethod
    def _rebuild_keyword_index(cls) -> None:
        """
        Builds {keyword_lower: ColumnDefinition} from all COLUMNS.
        Must be called after load_dynamic_columns() to include config-defined columns.
        """
        cls._KEYWORD_INDEX = {}
        for col_def in cls.COLUMNS.values():
            for keyword in col_def.keywords:
                # First keyword wins (hardcoded takes priority since loaded first)
                if keyword not in cls._KEYWORD_INDEX:
                    cls._KEYWORD_INDEX[keyword] = col_def
        logger.info(f"[BlueprintSchema] Keyword index built: {len(cls._KEYWORD_INDEX)} entries.")

    @classmethod
    def get_format_for_id(cls, col_id: str) -> str:
        """Get the defined Excel format for a column ID."""
        if col_id in cls.COLUMNS:
            return cls.COLUMNS[col_id].excel_format
        return "@"

    @classmethod
    def load_dynamic_columns(cls, mapping_config: Dict) -> None:
        """
        Merge column definitions from the provided mapping_config dict into COLUMNS.
        This accepts a plain dict (caller's responsibility to get it from DB/file).
        """
        # Reset COLUMNS to the base set
        cls.COLUMNS = dict(cls._BASE_COLUMNS)

        # Fallback to disk config file if mapping_config is empty or lacks shipping_header_map
        if not mapping_config or "shipping_header_map" not in mapping_config:
            try:
                from core.system_config import sys_config
                disk_path = sys_config.mapping_config_path
                if disk_path.exists():
                    with open(disk_path, 'r', encoding='utf-8') as f:
                        disk_config = json.load(f)
                        if mapping_config:
                            # Keep whatever original keys were passed in (e.g. footer_label_mappings)
                            disk_config.update(mapping_config)
                        mapping_config = disk_config
            except Exception as e:
                logger.exception("[BlueprintSchema] Failed to load fallback config from disk")
                raise e

        try:
            col_defs = mapping_config.get("shipping_header_map", {})
            loaded = 0
            for col_id, props in col_defs.items():
                # Skip metadata keys like 'comment'
                if not isinstance(props, dict):
                    continue
                
                keywords = [kw.lower() for kw in props.get("keywords", [])]
                excel_format = props.get("format", "@")
                width = float(props.get("width", 15.0))
                
                # Merge dynamic keywords/formatting into existing base columns
                if col_id in cls.COLUMNS:
                    cls.COLUMNS[col_id].keywords = keywords
                    if excel_format != "@":
                        cls.COLUMNS[col_id].excel_format = excel_format
                    if width != 15.0:
                        cls.COLUMNS[col_id].width = width
                else:
                    cls.COLUMNS[col_id] = ColumnDefinition(
                        id=col_id,
                        keywords=keywords,
                        excel_format=excel_format,
                        width=width
                    )
                    loaded += 1

            if loaded:
                logger.info(f"[BlueprintSchema] Loaded {loaded} new column definition(s) from mapping_config.")

        except Exception as e:
            logger.exception("[BlueprintSchema] Failed to load shipping_header_map from mapping_config")
            raise e

        # Load Sheet Classification Rules dynamically
        aggregation_sheets = mapping_config.get("aggregation_sheets")
        if aggregation_sheets:
            cls.AGGREGATION_SHEETS = set(aggregation_sheets)

        processed_tables_sheets = mapping_config.get("processed_tables_sheets")
        if processed_tables_sheets:
            cls.PROCESSED_TABLES_SHEETS = set(processed_tables_sheets)

        cls.ALLOWED_SEARCH_SHEETS = cls.AGGREGATION_SHEETS | cls.PROCESSED_TABLES_SHEETS

        # Rebuild the keyword index after loading dynamic columns
        cls._rebuild_keyword_index()

    # 4. Standard Row Heights (Fallback)
    # derived from JF_v2_bundle_config.json
    STANDARD_ROW_HEIGHTS: Dict[str, Dict[str, float]] = {
        "dataset_default": { # Fallback
            "header": 30.0,
            "data": 27.0,
            "footer": 30.0
        },
        "aggregation": { # Invoice, Contract
            "header": 35.0,
            "data": 35.0,
            "footer": 35.0
        },
        "processed_tables_multi": { # Packing List
            "header": 27.0,
            "data": 27.0,
            "footer": 27.0
        }
    }


# Initialize COLUMNS to base columns
BlueprintSchema.COLUMNS = dict(BlueprintSchema._BASE_COLUMNS)

# Load configuration from disk to initialize sheet categories and baseline columns at import time safely
try:
    from core.system_config import sys_config
    disk_path = sys_config.mapping_config_path
    if disk_path.exists():
        with open(disk_path, 'r', encoding='utf-8') as f:
            disk_config = json.load(f)
            BlueprintSchema.load_dynamic_columns(disk_config)
except Exception as e:
    # Build keyword index fallback if disk config load fails
    BlueprintSchema._rebuild_keyword_index()
