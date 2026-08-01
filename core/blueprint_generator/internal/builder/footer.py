"""
Footer Builder - Builds footer section config for blueprint generation.

Extracts footer row construction logic from ConfigBuilder, including:
- Main total row with SUM formulas
- Addon rows (leather summary, etc.) via AddonRegistry
"""

import logging
from typing import Dict, List, Any, Set

from ..addons import AddonRegistry
from ..scanner import SheetAnalysis

logger = logging.getLogger(__name__)


def build_footer(sheet: SheetAnalysis) -> Dict[str, Any]:
    """
    Build footer section for a sheet using the declarative rows layout schema.
    """
    rows = []
    
    sheet_col_ids = []
    for c in sheet.columns:
        sheet_col_ids.append(c.id)
        sheet_col_ids.extend(child.id for child in c.children)

    # 1. Main Footer Row
    main_footer_row = []
    if sheet.footer_info:
        # TOTAL label
        total_col = sheet.footer_info.total_text_col_id or "col_no"
        total_text = sheet.footer_info.total_text or "TOTAL:"
        total_cell = {
            "col_id": total_col,
            "value": total_text,
            "style_context": "footer"
        }
        if sheet.footer_info.merge_curr_colspan > 1:
            total_cell["colspan"] = sheet.footer_info.merge_curr_colspan
        main_footer_row.append(total_cell)

        # Pallet count
        if sheet.footer_info.pallet_count_col_id:
            main_footer_row.append({
                "col_id": sheet.footer_info.pallet_count_col_id,
                "value": "{col_pallet_count} PALLET{multiple}",
                "style_context": "footer"
            })
    else:
        main_footer_row.append({
            "col_id": "col_desc" if "col_desc" in sheet_col_ids else "col_po",
            "value": "TOTAL:",
            "style_context": "footer"
        })

    # Add SUM formulas for default numeric columns that exist in the sheet
    default_numeric_cols = ["col_qty_pcs", "col_qty_sf", "col_amount", "col_net", "col_gross", "col_cbm", "col_sqm"]
    for col_id in default_numeric_cols:
        if col_id in sheet_col_ids:
            main_footer_row.append({
                "col_id": col_id,
                "formula": "SUM",
                "target_section": "data",
                "style_context": "footer"
            })


    if main_footer_row:
        rows.append(main_footer_row)

    return {
        "rows": rows
    }


