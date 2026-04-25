"""
Blueprint Rules - Single Source of Truth for Invoice Generation Business Logic.

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

class BlueprintRules:
    """Central registry of business rules and definitions."""

    # 1. Sheet Classification Rules
    # 1. Sheet Classification Rules
    # Aggregation: Single table, usually financials (Invoice, Contract, Summary Packing List)
    AGGREGATION_SHEETS: Set[str] = {"invoice", "contract", "inv", "commercial", "shipping", "bill", "summary_packing_list"}
    # Processed Tables: Multiple tables or line items (Packing List, Details)
    PROCESSED_TABLES_SHEETS: Set[str] = {"packing list", "packing", "pl", "detail", "content", "weight", "detail_packing_list"}
    # Allowed search sheets: Union of all recognized sheet types that should be scanned
    ALLOWED_SEARCH_SHEETS: Set[str] = AGGREGATION_SHEETS | PROCESSED_TABLES_SHEETS

    # 2. Column Definitions
    # These replace the hardcoded HEADER_MAPPINGS in excel_scanner.py
    # and the hardcoded format checks.
    COLUMNS: Dict[str, ColumnDefinition] = {
        "col_static": ColumnDefinition(
            id="col_static", 
            keywords=["mark & n", "mark & no", "mark & nº", "mark and no"], 
            excel_format="@",
            width=24.71
        ),
        "col_po": ColumnDefinition(
            id="col_po", 
            keywords=["p.o", "po", "p.o.", "p.o nº", "p.o. nº", "po no"], 
            excel_format="@",
            width=28.0
        ),
        "col_item": ColumnDefinition(
            id="col_item", 
            keywords=["item", "item n", "item no", "item nº", "item. no", "item. nº"], 
            excel_format="@",
            width=22.14
        ),
        "col_desc": ColumnDefinition(
            id="col_desc", 
            keywords=["description"], 
            excel_format="@",
            width=26.0
        ),
        # Quantity Headers
        "col_qty_header": ColumnDefinition(
            id="col_qty_header", 
            keywords=["quantity"], 
            excel_format="@",
            width=15.0
        ),
        "col_qty_pcs": ColumnDefinition(
            id="col_qty_pcs", 
            keywords=["pcs"], 
            excel_format="#,##0",
            width=15.0
        ),
        "col_qty_sf": ColumnDefinition(
            id="col_qty_sf", 
            keywords=["sf", "sqft", "quantity(sf)"], 
            excel_format="#,##0.00",
            width=15.0
        ),
        "col_unit_sf": ColumnDefinition(
            id="col_unit_sf",
            keywords=["unit/sf", "unit sf", "unit(sf)"],
            excel_format="#,##0.00",
            width=15.0
        ),
        # Financials
        "col_unit_price": ColumnDefinition(
            id="col_unit_price", 
            keywords=["unit price", "price", "unit price (usd)", "unit price(usd)"], 
            excel_format="#,##0.00",
            width=15.0
        ),
        "col_amount": ColumnDefinition(
            id="col_amount", 
            keywords=["amount", "total", "value", "amount (usd)", "total value(usd)"], 
            excel_format="#,##0.00",
            width=18.0
        ),
        # Weights & Measures
        "col_net": ColumnDefinition(
            id="col_net", 
            keywords=["n.w", "net", "nw", "net weight", "n.w (kgs)"], 
            excel_format="#,##0.00"
        ),
        "col_gross": ColumnDefinition(
            id="col_gross", 
            keywords=["g.w", "gross", "gw", "gross weight", "g.w (kgs)"], 
            excel_format="#,##0.00"
        ),
        "col_cbm": ColumnDefinition(
            id="col_cbm", 
            keywords=["cbm", "m3"], 
            excel_format="0.00"
        ),
        # Others
        "col_no": ColumnDefinition(
            id="col_no", 
            keywords=["no", "no."], 
            excel_format="@"
        ),
        "col_pallet_count": ColumnDefinition(
            id="col_pallet_count", 
            keywords=["pallet", "plt", "pallet no", "pallet no.", "plt no", "plt no.", "pallet #", "pallet number", "pallet no.#", "pallet no. #"], 
            excel_format="@"
        ),
        "col_dc": ColumnDefinition(
            id="col_dc",
            keywords=["dc", "dc no", "dc #"],
            excel_format="@",
            width=12.0
        ),
        "col_container_no": ColumnDefinition(
            id="col_container_no",
            keywords=["container no", "container no.", "container #"],
            excel_format="@",
            width=18.0
        ),
        "col_remarks": ColumnDefinition(
            id="col_remarks",
            keywords=["remark", "remarks", "note", "notes", "comment", "comments"],
            excel_format="@",
            width=20.0
        ),
        "col_sqm": ColumnDefinition(
            id="col_sqm",
            keywords=["sqm", "m²", "square meter"],
            excel_format="#,##0.00",
            width=16.0
        ),
        "col_hs_code": ColumnDefinition(
            id="col_hs_code",
            keywords=["hs code", "h.s. code", "hscode"],
            excel_format="@",
            width=15.0
        ),
    }

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
            tick("rules.get_column_by_keyword", sub="index_hits")
            return result
        
        tick("rules.get_column_by_keyword", sub="index_misses")
        
        # 2. Smart fallback for HS Code (regex-like: both 'hs' and 'code' present)
        header_clean = re.sub(r'[^a-z0-9]', '', header_lower)
        if 'hs' in header_clean and 'code' in header_clean:
            return cls.COLUMNS.get("col_hs_code")
                    
        return None

    @classmethod
    def _rebuild_keyword_index(cls) -> None:
        """
        Builds {keyword_lower: ColumnDefinition} from all COLUMNS.
        Must be called after _load_from_config() to include JSON-defined columns.
        """
        cls._KEYWORD_INDEX = {}
        for col_def in cls.COLUMNS.values():
            for keyword in col_def.keywords:
                # First keyword wins (hardcoded takes priority since loaded first)
                if keyword not in cls._KEYWORD_INDEX:
                    cls._KEYWORD_INDEX[keyword] = col_def
        logger.info(f"[BlueprintRules] Keyword index built: {len(cls._KEYWORD_INDEX)} entries.")

    @classmethod
    def get_format_for_id(cls, col_id: str) -> str:
        """Get the defined Excel format for a column ID."""
        if col_id in cls.COLUMNS:
            return cls.COLUMNS[col_id].excel_format
        return "@"

    @classmethod
    def _load_from_config(cls) -> None:
        """
        Merge column definitions from mapping_config.json into COLUMNS.

        Reads the 'shipping_header_map' section from the JSON config and
        creates ColumnDefinition entries for any ID not already hardcoded
        here. This makes mapping_config.json the source of truth for
        new/custom columns without requiring Python code changes.

        Hardcoded COLUMNS always take priority over JSON definitions.
        """
        try:
            from core.system_config import sys_config
            json_path = sys_config.mapping_config_path

            if not json_path.exists():
                logger.warning(f"mapping_config.json not found at {json_path}. Skipping dynamic column load.")
                return

            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            col_defs = data.get("shipping_header_map", {})
            loaded = 0
            for col_id, props in col_defs.items():
                # Skip metadata keys like 'comment'
                if not isinstance(props, dict):
                    continue
                # Never override a hardcoded definition
                if col_id in cls.COLUMNS:
                    continue
                keywords = [kw.lower() for kw in props.get("keywords", [])]
                excel_format = props.get("format", "@")
                width = float(props.get("width", 15.0))
                cls.COLUMNS[col_id] = ColumnDefinition(
                    id=col_id,
                    keywords=keywords,
                    excel_format=excel_format,
                    width=width
                )
                loaded += 1

            if loaded:
                logger.info(f"[BlueprintRules] Loaded {loaded} column definition(s) from mapping_config.json.")

        except Exception as e:
            logger.error(f"[BlueprintRules] Failed to load shipping_header_map from mapping_config.json: {e}")

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


# Load JSON-defined columns once at module import time.
BlueprintRules._load_from_config()
# Build keyword index AFTER all columns (hardcoded + JSON) are loaded.
BlueprintRules._rebuild_keyword_index()
