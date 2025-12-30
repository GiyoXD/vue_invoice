"""
Blueprint Rules - Single Source of Truth for Invoice Generation Business Logic.

This module defines the business rules for:
1. Column Identification (keywords -> column definition)
2. Sheet Classification (name -> data source)
3. Column Formatting (id -> excel number format)
"""

from dataclasses import dataclass, field
from typing import List, Dict, Set, Optional

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
    # Aggregation: Single table, usually financials (Invoice, Contract)
    AGGREGATION_SHEETS: Set[str] = {"invoice", "contract", "inv", "commercial", "shipping", "bill"}
    # Processed Tables: Multiple tables or line items (Packing List, Details)
    PROCESSED_TABLES_SHEETS: Set[str] = {"packing list", "packing", "pl", "detail", "content", "weight"}

    # 2. Column Definitions
    # These replace the hardcoded HEADER_MAPPINGS in excel_scanner.py
    # and the hardcoded format checks.
    COLUMNS: Dict[str, ColumnDefinition] = {
        "col_static": ColumnDefinition(
            id="col_static", 
            keywords=["mark", "mark & n", "mark & no", "mark & nº", "mark and no"], 
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
            keywords=["description", "desc"], 
            excel_format="@",
            width=26.0
        ),
        # Quantity Headers
        "col_qty_header": ColumnDefinition(
            id="col_qty_header", 
            keywords=["quantity", "qty"], 
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
            keywords=["sf", "sqft", "ft2", "quantity(sf)"], 
            excel_format="#,##0.00",
            width=15.0
        ),
        # Financials
        "col_unit_price": ColumnDefinition(
            id="col_unit_price", 
            keywords=["unit price", "unit", "price", "unit price (usd)", "unit price(usd)"], 
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
        "col_pallet": ColumnDefinition(
            id="col_pallet", 
            keywords=["pallet", "plt"], 
            excel_format="@"
        ),
    }

    @classmethod
    def get_column_by_keyword(cls, header_text: str) -> Optional[ColumnDefinition]:
        """
        Identify a column definition based on header text.
        Returns the Best Match (or None).
        """
        if not header_text:
            return None
            
        header_lower = header_text.lower().strip()
        
        # Check all definitions
        for col_def in cls.COLUMNS.values():
            for keyword in col_def.keywords:
                # Exact match or word boundary check could be used,
                # but legacy logic used "if keyword in header_lower".
                # We preserve that behavior for now.
                # STRICT matching requested by user.
                # Previously: if keyword in header_lower:
                if keyword == header_lower:
                    return col_def
                    
        return None

    @classmethod
    def get_format_for_id(cls, col_id: str) -> str:
        """Get the defined Excel format for a column ID."""
        if col_id in cls.COLUMNS:
            return cls.COLUMNS[col_id].excel_format
        return "@"

    # 3. Default Footer Sums
    # Defines which columns should be summed in the footer by default
    DEFAULT_FOOTER_SUMS: Dict[str, List[str]] = {
        "aggregation": ["col_qty_sf", "col_amount"],
        "processed_tables_multi": ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm"]
    }

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
