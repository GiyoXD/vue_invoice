import ast
import re
import logging
from typing import Any, Dict, List, Tuple, Union, Optional
from core.invoice_generator.data.data_preparer import _to_numeric

logger = logging.getLogger(__name__)

def extract_table_data(data_source: Any, data_source_type: str) -> Any:
    """
    Extract data for the specific table being processed.
    
    For multi-table data sources, this extracts the subset for table_key.
    For single-table sources, returns the full data source.
    """
    if data_source is None:
        return None
    
    # For processed_tables_multi, BuilderConfigResolver already extracted the table
    if data_source_type in ['processed_tables', 'processed_tables_multi']:
        return data_source
    
    # For other types like aggregation, return as-is
    # Check for stringified tuple keys (JSON artifact) and convert back to tuples
    if isinstance(data_source, dict):
        new_data = {}
        for k, v in data_source.items():
            if isinstance(k, str) and k.startswith('(') and k.endswith(')'):
                try:
                    # Clean up Decimal wrappers for literal_eval: "Decimal('1.2')" -> "1.2"
                    clean_k = re.sub(r"Decimal\((['\"])(.*?)\1\)", r"\2", k)
                    new_key = ast.literal_eval(clean_k)
                    new_data[new_key] = v
                except (ValueError, SyntaxError):
                    new_data[k] = v
            else:
                new_data[k] = v
        return new_data

    return data_source


def merge_static_content(
    data_rows: List[Dict[str, Any]],
    static_content: Dict[str, Any],
    dynamic_mapping_rules: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool
) -> None:
    """
    Merge static content into the first N data rows.
    Extends data_rows if there are more static values than data rows.
    """
    if not static_content or 'col_static' not in static_content:
        return

    static_values = static_content['col_static']
    static_col_id = 'col_static'
    
    if static_values and len(data_rows) > 0:
        # Resolve {col_desc_fallback} placeholder for dynamic static content
        desc_fallback_str = ""
        for rule_key, rule in dynamic_mapping_rules.items():
            if rule_key == 'col_desc' and isinstance(rule, dict):
                fallback_cfg = rule.get('fallback')
                if isinstance(fallback_cfg, dict):
                    if DAF_mode and 'daf' in fallback_cfg:
                        desc_fallback_str = fallback_cfg['daf']
                    elif custom_mode and 'custom' in fallback_cfg:
                        desc_fallback_str = fallback_cfg['custom']
                    elif 'standard' in fallback_cfg:
                        desc_fallback_str = fallback_cfg['standard']
                elif fallback_cfg is not None:
                    desc_fallback_str = str(fallback_cfg)
                break
        
        num_static_values = len(static_values)
        
        # Extend data_rows if we have more static values than data rows
        while len(data_rows) < num_static_values:
            data_rows.append({})

        for i, static_value in enumerate(static_values):
            if isinstance(static_value, str) and "{col_desc_fallback}" in static_value:
                static_value = static_value.replace("{col_desc_fallback}", str(desc_fallback_str))
                
            data_rows[i][static_col_id] = static_value
        
        logger.info(f"Merged {num_static_values} static values into {len(data_rows)} data rows")


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


def format_pallet_counts(
    data_rows: List[Dict[str, Any]],
    num_data_rows: int,
    pallet_col_id: Optional[str],
    footer_data: Dict[str, Any],
    table_key: Optional[Any]
) -> None:
    """
    Formats raw binary pallet counts (1/0) into display values (e.g., '1-5', '2-5')
    and carries values forward for proper vertical cell merging.
    Modifies data_rows in-place.
    """
    if num_data_rows <= 0 or not pallet_col_id:
        return
        
    global_total_pallets = 0
    starting_pallet_order = 0
    
    if footer_data:
        if 'grand_total' in footer_data and 'col_pallet_count' in footer_data['grand_total']:
            global_total_pallets = int(footer_data['grand_total']['col_pallet_count'])
            
        if 'table_totals' in footer_data and table_key is not None:
            table_totals = footer_data['table_totals']
            if isinstance(table_totals, list):
                tbl_idx = 0
                if str(table_key).isdigit():
                    tbl_idx = int(table_key)
                
                for i in range(min(tbl_idx, len(table_totals))):
                    tbl_footer = table_totals[i]
                    if 'col_pallet_count' in tbl_footer:
                        starting_pallet_order += int(tbl_footer['col_pallet_count'])

    total_pallets_to_display = global_total_pallets
    pallet_order = starting_pallet_order
    carry_value = 0
    
    for row in data_rows[:num_data_rows]:
        val = _to_numeric(row.get(pallet_col_id, 0))
        
        if val == 1:
            pallet_order += 1
            formatted_val = f"{pallet_order}-{total_pallets_to_display}"
            row[pallet_col_id] = formatted_val
            carry_value = formatted_val
        else:
            row[pallet_col_id] = carry_value if carry_value != 0 else 0
