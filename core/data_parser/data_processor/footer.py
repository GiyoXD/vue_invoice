import logging
from typing import List, Dict, Any

logger = logging.getLogger(__name__)


def calculate_all_footers(
    processed_tables: List[List[Dict[str, Any]]],
    merged_processed_data: List[Dict[str, Any]],
    normal_aggregate_per_po: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Orchestrates the calculation of all footer and addon data from processed tables.

    Returns a dictionary containing:
        - table_footer_data: List of footer totals per table.
        - grand_total_footer: Overall grand totals across all tables.
        - leather_summary: Summary of COW and BUFFALO leather types.
        - weight_summary_addon: Net and gross weight sums as floats.
    """
    from .aggregation import (
        calculate_leather_summary,
        calculate_footer_totals,
    )

    logger.info("--- Calculating Add-on Data ---")

    # Calculate leather summary (BUFFALO vs COW) across all tables
    # Use the normal_aggregate_per_po data so we get integer pallet counts and correct sums
    leather_summary = calculate_leather_summary(normal_aggregate_per_po)
    logger.info(f"Leather Summary: {leather_summary}")

    # --- Calculate Footer Data ---
    logger.info("--- Calculating Footer Data ---")

    # Calculate per-table totals
    table_footer_data = []
    for table_index, table_data in enumerate(processed_tables):
        table_id = str(table_index + 1)
        if isinstance(table_data, list):
            footer_totals = calculate_footer_totals(table_data)

            # If there's only one table, it gets all the pallets. If multiple, we'd need to distribute, 
            # but for now we'll rely on the parser to not double count.

            table_footer_data.append(footer_totals)
            logger.info(f"Table {table_id} Footer: {footer_totals}")

    # Calculate grand total (merged across all tables)
    grand_total_footer = calculate_footer_totals(merged_processed_data)
    logger.info(f"Grand Total Footer: {grand_total_footer}")

    # Calculate weight summary across all tables by reusing grand_total_footer
    weight_summary_addon = {
        'net': float(grand_total_footer.get('col_net', 0.0)),
        'gross': float(grand_total_footer.get('col_gross', 0.0))
    }
    logger.info(f"Weight Summary Addon: {weight_summary_addon}")

    return {
        "table_footer_data": table_footer_data,
        "grand_total_footer": grand_total_footer,
        "leather_summary": leather_summary,
        "weight_summary_addon": weight_summary_addon
    }
