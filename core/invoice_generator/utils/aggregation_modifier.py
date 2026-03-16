import logging
from typing import Dict, Any, List

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
                row["col_amount"] = str(round(current_val + final_chunk, 3))
            else:
                row["col_amount"] = str(round(current_val + per_row_adj, 3))
                total_added += per_row_adj

            if "col_qty_sf" in row and "col_unit_price" in row:
                try:
                    qty = float(row["col_qty_sf"])
                    if qty > 0:
                        row["col_unit_price"] = str(round(float(row["col_amount"]) / qty, 4))
                except Exception:
                    pass
        except (ValueError, TypeError):
            continue


def apply_aggregation_adjustment(json_data: Dict[str, Any], price_adjustments: List[List[Any]]) -> Dict[str, Any]:
    """
    Distributes multiple adjustment values across the 'col_amount' of all valid rows.
    Also updates the footer grand totals and stores the adjustment list in json_data.
    
    Args:
        json_data: The fully parsed JSON data dictionary for an invoice
        price_adjustments: A list of [description, value] pairs.
        
    Returns:
        The modified JSON data dictionary
    """
    if not price_adjustments:
        return json_data

    # Calculate total adjustment value
    total_adjustment = 0.0
    valid_adjustments = []
    
    for adj in price_adjustments:
        if isinstance(adj, list) and len(adj) >= 2:
            try:
                desc = str(adj[0])
                val = float(adj[1])
                if val != 0:
                    total_adjustment += val
                    valid_adjustments.append([desc, val])
            except (ValueError, TypeError):
                logger.warning(f"Skipping invalid adjustment entry: {adj}")

    # Store the list of adjustments for reporting/templates
    json_data["price_adjustment"] = valid_adjustments

    if total_adjustment == 0:
        return json_data

    # 1. Modify all tables in single_table (Aggregated views)
    if "single_table" in json_data:
        single_table = json_data["single_table"]
        for key, value in single_table.items():
            if not isinstance(value, list) or not value:
                continue
            if not isinstance(value[0], dict) or "col_amount" not in value[0]:
                continue
            _apply_to_row_list(value, total_adjustment)

    # 3. Update Footer Totals
    if "footer_data" in json_data:
        footer = json_data["footer_data"]
        
        # Grand Total
        if "grand_total" in footer and "col_amount" in footer["grand_total"]:
            try:
                curr_tot = float(footer["grand_total"]["col_amount"])
                footer["grand_total"]["col_amount"] = str(round(curr_tot + total_adjustment, 3))
            except (ValueError, TypeError):
                pass
                
        # Table Totals
        if "table_totals" in footer:
            for table_total in footer["table_totals"]:
                if "col_amount" in table_total:
                    try:
                        curr_tot = float(table_total["col_amount"])
                        table_total["col_amount"] = str(round(curr_tot + total_adjustment, 3))
                        # Note: If there are multiple tables, distributing the full adjustment 
                        # to EVERY table total might be incorrect depending on user intent.
                        # However, the previous logic only hit the first table.
                        # For now, we update the first table to match grand total if that's the primary one.
                        break 
                    except (ValueError, TypeError):
                        pass

    logger.info(f"Successfully applied {len(valid_adjustments)} price adjustments (Total: {total_adjustment})")
    return json_data
