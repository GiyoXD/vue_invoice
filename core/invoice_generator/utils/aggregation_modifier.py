import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)

def _apply_to_row_list(rows, adjustment_value: float) -> None:
    """
    Helper: distribute adjustment_value across a list of row dicts
    that contain 'col_amount', updating 'col_unit_price' when possible.
    """
    if not rows:
        return

    per_row_adj = round(adjustment_value / len(rows), 3)
    total_added = 0.0

    for i, row in enumerate(rows):
        if "col_amount" not in row:
            continue
        try:
            current_val = float(row["col_amount"])

            if i == len(rows) - 1:
                # Last row takes remainder to avoid rounding drift
                final_chunk = adjustment_value - total_added
                row["col_amount"] = round(current_val + final_chunk, 3)
            else:
                row["col_amount"] = round(current_val + per_row_adj, 3)
                total_added += per_row_adj

            if "col_qty_sf" in row and "col_unit_price" in row:
                try:
                    qty = float(row["col_qty_sf"])
                    if qty > 0:
                        row["col_unit_price"] = round(float(row["col_amount"]) / qty, 4)
                except Exception:
                    pass
        except (ValueError, TypeError):
            continue


def apply_aggregation_adjustment(json_data: Dict[str, Any], adjustment_value: float) -> Dict[str, Any]:
    """
    Distributes an adjustment value (positive or negative) evenly across the 
    'col_amount' of all valid rows in the aggregation lists.
    Also updates the footer grand totals so the final math matches.
    
    Args:
        json_data: The fully parsed JSON data dictionary for an invoice
        adjustment_value: The numerical amount to distribute (e.g. shipping cost, discount)
        
    Returns:
        The modified JSON data dictionary
    """
    if not adjustment_value or adjustment_value == 0:
        return json_data

    # 1. Modify all single_table aggregations (any list whose key starts with 'aggregation')
    if "single_table" in json_data:
        single_table = json_data["single_table"]
        for key, value in single_table.items():
            if not isinstance(value, list):
                continue
            if not key.startswith("aggregation"):
                continue
            if not value:
                continue
            if not isinstance(value[0], dict):
                continue
            _apply_to_row_list(value, adjustment_value)

    # 2. Modify Multi-Table (Packing List style) - usually we just apply to the first table chunk
    if "multi_table" in json_data and len(json_data["multi_table"]) > 0:
        # For multi-table, flatten all rows to count them
        all_rows = []
        for table in json_data["multi_table"]:
            all_rows.extend(table)
            
        if all_rows and len(all_rows) > 0:
            _apply_to_row_list(all_rows, adjustment_value)

    # 4. Update Footer Totals so the math matches at the bottom
    if "footer_data" in json_data:
        footer = json_data["footer_data"]
        
        # Grand Total
        if "grand_total" in footer and "col_amount" in footer["grand_total"]:
            try:
                curr_tot = float(footer["grand_total"]["col_amount"])
                footer["grand_total"]["col_amount"] = str(round(curr_tot + adjustment_value, 3))
            except (ValueError, TypeError):
                pass
                
        # Table Totals (just hit the first one if it exists)
        if "table_totals" in footer and len(footer["table_totals"]) > 0:
            if "col_amount" in footer["table_totals"][0]:
                try:
                    curr_tot = float(footer["table_totals"][0]["col_amount"])
                    footer["table_totals"][0]["col_amount"] = str(round(curr_tot + adjustment_value, 3))
                except (ValueError, TypeError):
                    pass

    logger.info(f"Successfully applied aggregation adjustment of {adjustment_value}")
    return json_data
