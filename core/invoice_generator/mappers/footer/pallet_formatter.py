import logging
from typing import Any, Dict, List, Optional
from core.invoice_generator.utils.math_utils import to_numeric

logger = logging.getLogger(__name__)


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
        val = to_numeric(row.get(pallet_col_id, 0))
        
        if val == 1:
            pallet_order += 1
            formatted_val = f"{pallet_order}-{total_pallets_to_display}"
            row[pallet_col_id] = formatted_val
            carry_value = formatted_val
        else:
            row[pallet_col_id] = carry_value if carry_value != 0 else 0
