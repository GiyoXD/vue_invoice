from typing import List, Dict, Any, Set
from core.blueprint_generator.internal.scanner.models.addons import BaseAddonFact, LeatherSummaryFact
from .base import BaseAddonBuilder

class LeatherAddonBuilder(BaseAddonBuilder):
    """Formats LeatherSummaryFact into JSON rows."""
    
    def build_rows(self, fact: BaseAddonFact, sheet_col_ids: Set[str]) -> List[List[Dict[str, Any]]]:
        assert isinstance(fact, LeatherSummaryFact), f"Expected LeatherSummaryFact, got {type(fact)}"
        
        leather_row = [

            {"col_id": fact.total_col_id, "value": fact.total_value, "style_context": "footer_addon", "addon_type": "leather", "leather_key": fact.leather_key},
            {"col_id": fact.label_col_id, "value": fact.label_value, "style_context": "footer_addon", "addon_type": "leather", "leather_key": fact.leather_key},
            {"col_id": "col_desc", "value": "{pallet_count} PALLET{multiple}", "style_context": "footer_addon", "addon_type": "leather", "leather_key": fact.leather_key}
        ]
        
        leather_cols = ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm", "col_sqm"]
        for col_id in leather_cols:
            if col_id in sheet_col_ids:
                leather_row.append({"col_id": col_id, "style_context": "footer_addon", "addon_type": "leather", "leather_key": fact.leather_key})
                
        return [leather_row]
