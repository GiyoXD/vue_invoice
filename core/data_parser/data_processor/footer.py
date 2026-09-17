import logging
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)


def calculate_all_footers(
    processed_tables: List[List[Dict[str, Any]]],
    merged_processed_data: List[Dict[str, Any]],
    normal_aggregate_per_po: List[Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None
) -> Dict[str, Any]:
    """
    Orchestrates the calculation of all footer and addon data from processed tables.

    Returns a dictionary containing:
        - table_footer_data: List of footer totals per table.
        - grand_total_footer: Overall grand totals across all tables.
        - leather_summary: Summary of COW and BUFFALO leather types.
        - weight_summary_addon: Net and gross weight sums as floats.
    """
    processed_tables = processed_tables or []
    merged_processed_data = merged_processed_data or []
    normal_aggregate_per_po = normal_aggregate_per_po or []

    from .single_table import (
        calculate_leather_summary,
        calculate_footer_totals,
    )

    logger.info("--- Calculating Add-on Data ---")

    # Calculate leather summary (BUFFALO vs COW) across all tables
    # Use the normal_aggregate_per_po data so we get integer pallet counts and correct sums
    leather_summary = calculate_leather_summary(normal_aggregate_per_po, decimals_map=decimals_map)
    logger.info(f"Leather Summary: {leather_summary}")

    # --- Calculate Footer Data ---
    logger.info("--- Calculating Footer Data ---")

    # Calculate per-table totals
    table_footer_data = []
    for table_index, table_data in enumerate(processed_tables):
        table_id = str(table_index + 1)
        if isinstance(table_data, list):
            footer_totals = calculate_footer_totals(table_data, decimals_map=decimals_map)

            # If there's only one table, it gets all the pallets. If multiple, we'd need to distribute, 
            # but for now we'll rely on the parser to not double count.

            table_footer_data.append(footer_totals)
            logger.info(f"Table {table_id} Footer: {footer_totals}")

    # Calculate grand total (merged across all tables)
    grand_total_footer = calculate_footer_totals(merged_processed_data, decimals_map=decimals_map)
    logger.info(f"Grand Total Footer: {grand_total_footer}")

    # Calculate weight summary across all tables by reusing grand_total_footer
    weight_summary_addon = {
        'net': float(grand_total_footer.get('col_net') or 0.0),
        'gross': float(grand_total_footer.get('col_gross') or 0.0)
    }
    logger.info(f"Weight Summary Addon: {weight_summary_addon}")

    return {
        "table_footer_data": table_footer_data,
        "grand_total_footer": grand_total_footer,
        "leather_summary": leather_summary,
        "weight_summary_addon": weight_summary_addon
    }


def format_footer_payload(footer_results: Dict[str, Any]) -> Dict[str, Any]:
    """
    Formats raw footer calculation results into the canonical footer_data schema
    expected in the exported JSON and invoice generator.

    Transforms:
        - table_footer_data -> table_totals
        - grand_total_footer -> grand_total
        - leather_summary (dict) -> leather_summary (list of records)
    """
    footer_results = footer_results or {}
    leather_summary_records = []
    raw_leather_summary = footer_results.get("leather_summary", {})
    if isinstance(raw_leather_summary, dict):
        for l_type, l_data in raw_leather_summary.items():
            if not isinstance(l_data, dict):
                continue
            pallet_count = l_data.get('col_pallet_count') or 0
            qty_sf = l_data.get('col_qty_sf') or 0.0
            if pallet_count == 0 and qty_sf == 0:
                continue

            p_cnt = int(pallet_count)
            record = {
                "leather_type": l_type,
                "col_pallet_count": p_cnt,
                "multiple": "S" if p_cnt != 1 else "",
                "col_qty_pcs": int(l_data.get('col_qty_pcs') or 0),
                "col_qty_sf": float(qty_sf),
                "col_net": float(l_data.get('col_net') or 0.0),
                "col_gross": float(l_data.get('col_gross') or 0.0),
                "col_cbm": float(l_data.get('col_cbm') or 0.0),
            }
            leather_summary_records.append(record)
    elif isinstance(raw_leather_summary, list):
        leather_summary_records = raw_leather_summary

    return {
        "table_totals": footer_results.get("table_totals", footer_results.get("table_footer_data", [])),
        "grand_total": footer_results.get("grand_total", footer_results.get("grand_total_footer", {})),
        "leather_summary": leather_summary_records,
    }

