from typing import List, Dict, Any, Set
from core.blueprint_generator.internal.scanner.models.addons import BaseAddonFact, LeatherSummaryFact
from .base import BaseAddonBuilder

class LeatherAddonBuilder(BaseAddonBuilder):
    """Formats LeatherSummaryFact into JSON rows."""
    
    def build_rows(self, fact: BaseAddonFact, sheet_col_ids: Set[str]) -> List[List[Dict[str, Any]]]:
        assert isinstance(fact, LeatherSummaryFact), f"Expected LeatherSummaryFact, got {type(fact)}"
        
        # Build dynamic label template from scanned text
        # e.g. "BUFFALO LEATHER" → "{leather_type} LEATHER"
        dynamic_label = fact.label_value.replace(fact.leather_key.upper(), "{leather_type}").replace(fact.leather_key.lower(), "{leather_type}")
        
        cells = [
            {"col_id": fact.total_col_id, "value": fact.total_value, "style_context": "summary"},
            {"col_id": fact.label_col_id, "value": dynamic_label, "style_context": "summary"},
        ]
        
        if "col_desc" in sheet_col_ids and "col_desc" not in (fact.total_col_id, fact.label_col_id):
            cells.append({"col_id": "col_desc", "value": "{col_pallet_count} PALLET{multiple}", "style_context": "summary"})

        row_dict = {
            "source_list": "leather_summary",
            "cells": cells
        }
        
        leather_cols = ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm", "col_sqm"]
        for col_id in leather_cols:
            if col_id in sheet_col_ids:
                row_dict["cells"].append({"col_id": col_id, "style_context": "summary"})
                
        return [row_dict]
