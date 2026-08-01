import logging
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)

def get_data_source_for_type(
    data_source_type: str, 
    invoice_data: Optional[Dict[str, Any]], 
    sheet_name: str
) -> Any:
    """
    Extract the appropriate data source from invoice_data based on type.
    
    The data_source_type is expected to be pre-resolved by the config layer
    (e.g. 'daf' when DAF mode is active), so this function
    only performs direct key lookups without mode-flag inspection.
    """
    if not invoice_data:
        return {}
    
    single_table_group = invoice_data.get('single_table', {})
    multi_table_group = invoice_data.get('multi_table', {})
    
    # --- Multi-table types ---
    if data_source_type in ['processed_tables_multi', 'processed_tables', 'detail_packing_list']:
        if multi_table_group:
            logger.debug(f"Resolver: Found multi_table data for '{data_source_type}'")
            return multi_table_group
    
    # --- Single-table: direct key lookup (mode already resolved by config) ---
    if data_source_type in single_table_group:
        logger.debug(f"Resolver: Found '{data_source_type}' in single_table")
        return single_table_group[data_source_type]
    
    # --- Special fallback: summary_packing_list → manifest_by_pallet_per_po ---
    if data_source_type == 'summary_packing_list':
        if 'manifest_by_pallet_per_po' in single_table_group:
            logger.info("Summary packing list requested: Using 'manifest_by_pallet_per_po' from single_table group")
            return single_table_group['manifest_by_pallet_per_po']
    
    logger.warning(f"Resolver: Data source '{data_source_type}' not found in single_table or multi_table groups.")
    return {}

