import logging
from typing import Dict, Any, List

logger = logging.getLogger(__name__)

def _is_discount_row(row: Dict[str, Any]) -> bool:
    """Check if a row is a discount row."""
    if row.get("is_discount"):
        return True
    col_level = str(row.get("col_level", ""))
    if "折扣" in col_level:
        return True
    try:
        if float(row.get("col_amount", 0)) <= 0:
            return True
    except (ValueError, TypeError):
        pass
    return False


def _is_ghost_row(row: Dict[str, Any]) -> bool:
    """Check if a row is an empty Buffalo ghost row (col_amount == 0 and col_qty_sf == 0)."""
    try:
        amt = float(row.get("col_amount", 0))
        qty = float(row.get("col_qty_sf", 0))
        return amt == 0 and qty == 0
    except (ValueError, TypeError):
        return False


def _apply_to_row_list(rows, adjustment_value: float) -> None:
    """
    Helper: distribute adjustment_value pro-rata across eligible rows
    that contain 'col_amount', updating 'col_unit_price' when possible.
    Discount rows and empty ghost rows are excluded.
    """
    if not rows or adjustment_value == 0:
        return

    # Filter out empty Buffalo ghost rows (col_amount == 0 and col_qty_sf == 0)
    candidate_rows = [r for r in rows if isinstance(r, dict) and "col_amount" in r and not _is_ghost_row(r)]
    if not candidate_rows:
        return

    # Filter eligible primary rows: rows that are NOT discount rows, and have float(row['col_amount']) > 0
    eligible_rows = []
    for r in candidate_rows:
        if _is_discount_row(r):
            continue
        try:
            if float(r["col_amount"]) > 0:
                eligible_rows.append(r)
        except (ValueError, TypeError):
            continue

    # If no eligible rows found, fall back to all rows with float(row['col_amount']) != 0
    if not eligible_rows:
        for r in candidate_rows:
            try:
                if float(r["col_amount"]) != 0:
                    eligible_rows.append(r)
            except (ValueError, TypeError):
                continue

    if not eligible_rows:
        return

    # Calculate total eligible amount
    sum_eligible = sum(float(r["col_amount"]) for r in eligible_rows)
    if sum_eligible <= 0:
        return

    total_added = 0.0
    n = len(eligible_rows)

    for i, row in enumerate(eligible_rows):
        current_val = float(row["col_amount"])
        if i == n - 1:
            # Last eligible row gets remainder to prevent rounding drift
            share = round(adjustment_value - total_added, 3)
        else:
            share = round(adjustment_value * (current_val / sum_eligible), 3)
            total_added += share

        row["col_amount"] = str(round(current_val + share, 3))

        qty = 0.0
        if "col_qty_sf" in row:
            try:
                qty = float(row["col_qty_sf"])
            except (ValueError, TypeError):
                qty = 0.0
        if qty == 0.0 and "col_net" in row:
            try:
                qty = float(row["col_net"])
            except (ValueError, TypeError):
                qty = 0.0
        if qty > 0 and "col_unit_price" in row:
            row["col_unit_price"] = str(round(float(row["col_amount"]) / qty, 4))


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
