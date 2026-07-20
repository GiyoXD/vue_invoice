import logging
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


def extract_summaries(
    data_source: Any,
    footer_data: Dict[str, Any],
    table_key: Optional[Any]
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], Optional[int]]:
    """
    Extract leather_summary, weight_summary, and pallet_summary_total.
    """
    leather_summary = None
    weight_summary = None
    pallet_summary_total = None
    
    if isinstance(data_source, dict):
        leather_summary = data_source.get('leather_summary')
        if not leather_summary and footer_data and 'add_ons' in footer_data:
            leather_summary = footer_data['add_ons'].get('leather_summary_addon')
        
        weight_summary = data_source.get('weight_summary')
        if not weight_summary and footer_data and 'add_ons' in footer_data:
            weight_summary = footer_data['add_ons'].get('weight_summary_addon')
    elif isinstance(data_source, list):
        if footer_data and 'add_ons' in footer_data:
            leather_summary = footer_data['add_ons'].get('leather_summary_addon')
            weight_summary = footer_data['add_ons'].get('weight_summary_addon')
            
    if footer_data:
        if table_key is None:
            if 'grand_total' in footer_data and 'col_pallet_count' in footer_data['grand_total']:
                pallet_summary_total = int(footer_data['grand_total']['col_pallet_count'])
        
        if pallet_summary_total is None and 'table_totals' in footer_data:
            table_totals = footer_data['table_totals']
            if isinstance(table_totals, list) and len(table_totals) > 0:
                tbl_idx = 0
                if table_key is not None and str(table_key).isdigit():
                    tbl_idx = int(table_key)
                
                if tbl_idx >= len(table_totals):
                    tbl_idx = 0
                
                tbl_footer = table_totals[tbl_idx]
                if 'col_pallet_count' in tbl_footer:
                    pallet_summary_total = int(tbl_footer['col_pallet_count'])
            elif isinstance(table_totals, dict):
                first_val = next(iter(table_totals.values()), {})
                if 'col_pallet_count' in first_val:
                    pallet_summary_total = int(first_val['col_pallet_count'])
                    
    if pallet_summary_total is None and isinstance(data_source, dict):
        pallet_summary_total = data_source.get('pallet_summary_total')
        if pallet_summary_total is not None:
            logger.warning(f"Using legacy pallet_summary_total from data_source: {pallet_summary_total}")
            
    return leather_summary, weight_summary, pallet_summary_total
