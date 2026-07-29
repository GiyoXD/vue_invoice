"""
Summary Builder - Builds summary section config for blueprint generation.
"""

import logging
from typing import Dict, List, Any

from ..addons import AddonRegistry
from ..scanner import SheetAnalysis

logger = logging.getLogger(__name__)


def build_summary(sheet: SheetAnalysis) -> Dict[str, Any]:
    """
    Build page-level summary section for a sheet.
    Contains weight summary rows and dynamic addon rows (leather summary, etc.).
    """
    rows = []

    sheet_col_ids = []
    for c in sheet.columns:
        sheet_col_ids.append(c.id)
        sheet_col_ids.extend(child.id for child in c.children)

    # 1. Weight summary declaration rows (Disabled by default - enable when invoice use case is found)
    # desc_col = "col_desc" if "col_desc" in sheet_col_ids else "col_po"
    # qty_col = "col_qty_pcs" if "col_qty_pcs" in sheet_col_ids else ("col_qty" if "col_qty" in sheet_col_ids else "col_item")

    # rows.append([
    #     {"col_id": desc_col, "value": "NET WEIGHT:", "style_context": "summary_value_only"},
    #     {"col_id": qty_col, "value": "{weight_net} KGS", "style_context": "summary_value_only"}
    # ])
    # rows.append([
    #     {"col_id": desc_col, "value": "GROSS WEIGHT:", "style_context": "summary_value_only"},
    #     {"col_id": qty_col, "value": "{weight_gross} KGS", "style_context": "summary_value_only"}
    # ])

    # 2. Addon rows (leather summary, etc.)
    addon_facts = sheet.static_content_hints.get("addon_facts", [])
    seen_source_lists = set()
    for fact in addon_facts:
        try:
            addon_builder = AddonRegistry.get_builder(fact)
            addon_rows = addon_builder.build_rows(fact, sheet_col_ids)
            for r in addon_rows:
                if isinstance(r, dict) and "source_list" in r:
                    src = r["source_list"]
                    if src in seen_source_lists:
                        continue
                    seen_source_lists.add(src)
                rows.append(r)
        except Exception as e:
            logger.warning(f"    Failed to build addon row for fact {getattr(fact, 'fact_type', 'unknown')}: {e}")

    return {
        "rows": rows
    }
