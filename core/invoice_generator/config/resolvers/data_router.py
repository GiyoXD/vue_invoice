import logging
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)

def get_data_source_for_type(
    data_source_type: str, 
    invoice_data: Optional[Dict[str, Any]], 
    sheet_name: str, 
    args: Any = None
) -> Any:
    """
    Extract the appropriate data source from invoice_data based on type.
    
    Refactored Logic (v2.2):
    1. STRUCTURED LOOKUP: Checks `single_table` for aggregation types or `multi_table` for granular ones.
    2. STRICT LOOKUP: Checks if 'data_source_type' exists as a direct key in invoice_data.
    3. LEGACY FALLBACK: Checks the hardcoded 'type_mapping' for backward compatibility.
    """
    if not invoice_data:
        return {}
        
    # --- PATH -1: Override generic aggregation for Summary Packing List ---
    # If the blueprint generator marked the sheet as generic "aggregation" but it is a
    # Summary Packing List, explicitly upgrade it so it routes to manifest_by_pallet_per_po.
    normalized_sheet = sheet_name.strip().lower()
    if data_source_type == 'aggregation' and normalized_sheet == 'summary packing list':
        logger.info("Auto-upgrading data_source_type from 'aggregation' to 'summary_packing_list' based on sheet name")
        data_source_type = 'summary_packing_list'
    
    # --- PATH 0: STRUCTURED LOOKUP (The "Newest Way" - v2.3) ---
    single_table_group = invoice_data.get('single_table', {})
    multi_table_group = invoice_data.get('multi_table', {})
    
    # Determine if we should look in single_table or multi_table
    if data_source_type in ['processed_tables_multi', 'processed_tables', 'detail_packing_list']:
        if multi_table_group:
            logger.debug(f"Smart Resolver: Found multi_table data for '{data_source_type}'")
            return multi_table_group
    else:
        # Single table types (aggregation, DAF_aggregation, custom_aggregation, summary_packing_list)
        # 1. Flag Checks: If DAF/Custom mode is on, look for suffixed keys first inside `single_table`.
        if args:
            if getattr(args, 'DAF', False):
                daf_key = f"{data_source_type}_DAF"
                if daf_key in single_table_group:
                    logger.debug(f"Smart Resolver: DAF Mode ON. Found variant '{daf_key}' in single_table")
                    return single_table_group[daf_key]
            
            if getattr(args, 'custom', False):
                custom_key = f"{data_source_type}_custom"
                if custom_key in single_table_group:
                    logger.debug(f"Smart Resolver: Custom Mode ON. Found variant '{custom_key}' in single_table")
                    return single_table_group[custom_key]
                    
        # 2. Base Lookup in single_table: If no flag match (or flags off), use exact key.
        if data_source_type in single_table_group:
            logger.debug(f"Smart Resolver: Found strict match for data_source '{data_source_type}' in single_table")
            return single_table_group[data_source_type]
        
        # 3. Handle Special Fallbacks strictly within single_table
        if data_source_type == 'summary_packing_list':
            if 'manifest_by_pallet_per_po' in single_table_group:
                logger.info("Summary packing list requested: Using 'manifest_by_pallet_per_po' from single_table group")
                return single_table_group['manifest_by_pallet_per_po']
                
        if data_source_type == 'aggregation' and args and getattr(args, 'custom', False):
            if 'aggregation_custom' in single_table_group:
                logger.info("Custom mode active: Using 'aggregation_custom' from single_table group")
                return single_table_group['aggregation_custom']
                
        if data_source_type in ['aggregation', 'DAF_aggregation'] and args and getattr(args, 'DAF', False):
            if 'aggregation_DAF' in single_table_group:
                logger.info("DAF mode active: Using 'aggregation_DAF' from single_table group")
                return single_table_group['aggregation_DAF']

    # --- If not found in structured paths, return empty ---
    logger.warning(f"Resolver: Data source '{data_source_type}' not found in structured single_table or multi_table groups.")
    return {}
