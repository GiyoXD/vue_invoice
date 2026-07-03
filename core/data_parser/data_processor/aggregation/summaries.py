import decimal
from typing import List, Dict, Any, Tuple
from .reducers import decimal_sum_reducer, int_sum_reducer


def calculate_weight_summary(processed_data: List[Dict[str, Any]]) -> Dict[str, decimal.Decimal]:
    """Calculates the weight summary (Net Weight and Gross Weight)."""
    summary = {'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0)}
    if not processed_data:
        return summary
        
    summary['col_net'] = decimal_sum_reducer([r.get('col_net') for r in processed_data])
    summary['col_gross'] = decimal_sum_reducer([r.get('col_gross') for r in processed_data])
    return summary


def calculate_pallet_summary(processed_data: List[Dict[str, Any]]) -> int:
    """Calculates the total pallet count for the table."""
    total_pallets = 0
    if not processed_data:
        return 0
        
    for row in processed_data:
        val = row.get('col_pallet_count')
        if val is not None:
            try:
                total_pallets += int(float(str(val)))
            except (ValueError, TypeError):
                pass
    return total_pallets


def calculate_footer_totals(processed_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Calculates totals for footer fields based on processed data."""
    totals = {
        "col_qty_pcs": 0,
        "col_qty_sf": decimal.Decimal(0),
        "col_net": decimal.Decimal(0),
        "col_gross": decimal.Decimal(0),
        "col_cbm": decimal.Decimal(0),
        "col_amount": decimal.Decimal(0),
        "col_pallet_count": 0
    }
    
    if not processed_data:
        return totals

    totals['col_qty_pcs'] = int_sum_reducer([r.get('col_qty_pcs') for r in processed_data])
    totals['col_qty_sf'] = decimal_sum_reducer([r.get('col_qty_sf') for r in processed_data])
    totals['col_net'] = decimal_sum_reducer([r.get('col_net') for r in processed_data])
    totals['col_gross'] = decimal_sum_reducer([r.get('col_gross') for r in processed_data])
    totals['col_cbm'] = decimal_sum_reducer([r.get('col_cbm') for r in processed_data])
    totals['col_amount'] = decimal_sum_reducer([r.get('col_amount') for r in processed_data])
    totals['col_pallet_count'] = int_sum_reducer([r.get('col_pallet_count') for r in processed_data])

    return totals


def format_aggregation_as_list(
    aggregation_map: Dict[Tuple, Dict[str, decimal.Decimal]],
    mode: str = 'standard'
) -> List[Dict[str, Any]]:
    """
    Converts the internal tuple-keyed aggregation map into a clean list of dictionaries
    suitable for JSON output. Removes tuple keys and uses 'col_' prefixed keys for all fields.
    """
    flattened_list = []
    
    for key_tuple, values in aggregation_map.items():
        row_dict = {}
        
        # Extract values from the tuple key based on mode
        if mode == 'standard':
            if len(key_tuple) >= 4:
                row_dict['col_po'] = str(key_tuple[0]) if key_tuple[0] is not None else ""
                row_dict['col_item'] = str(key_tuple[1]) if key_tuple[1] is not None else ""
                row_dict['col_unit_price'] = str(key_tuple[2]) if key_tuple[2] is not None else ""
                row_dict['col_desc'] = str(key_tuple[3]) if key_tuple[3] is not None else ""
            else:
                row_dict['col_po'] = str(key_tuple[0]) if len(key_tuple) > 0 else ""
                row_dict['col_item'] = str(key_tuple[1]) if len(key_tuple) > 1 else ""
                
        elif mode == 'custom':
            if len(key_tuple) >= 4:
                row_dict['col_po'] = str(key_tuple[0]) if key_tuple[0] is not None else ""
                row_dict['col_item'] = str(key_tuple[1]) if key_tuple[1] is not None else ""
                row_dict['col_desc'] = str(key_tuple[3]) if key_tuple[3] is not None else ""
            else:
                row_dict['col_po'] = str(key_tuple[0]) if len(key_tuple) > 0 else ""
                row_dict['col_item'] = str(key_tuple[1]) if len(key_tuple) > 1 else ""

        row_dict.update(values)
        flattened_list.append(row_dict)
        
    return flattened_list
