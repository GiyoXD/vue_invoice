from typing import List, Dict, Any, Optional

from .engine import Aggregator, coerce_input_row
from .reducers import (
    int_sum_reducer,
    decimal_sum_reducer,
    first_non_empty_reducer,
    first_value_reducer,
)
from .footers import calculate_footer_totals
from .serializer import format_aggregation_as_list
from .aggregations import (
    quantize_record,
    finalize_standard_aggregation,
    finalize_item_merged_aggregation,
    aggregate_standard_by_po_item_price,
    aggregate_custom_by_po_item,
    aggregate_by_item_merge_po,
    calculate_leather_summary,
    aggregate_per_po_with_pallets,
    execute_strategy,
)
from .daf import (
    perform_DAF_compounding,
    DAFCompoundingResult,
    FinalDAFResultType,
)
from .schema import (
    INVOICE_LINE_ITEM_COLUMNS,
    INVOICE_FOOTER_COLUMNS,
)


def build_single_table(
    data: List[Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None,
    chunk_size: int = 2,
    item_separator: str = "/",
    line_separator: str = "\n",
) -> Dict[str, Any]:
    data = [coerce_input_row(r) for r in data if r]
    std_map = {}
    cust_map = {}
    aggregate_standard_by_po_item_price(data, std_map, decimals_map=decimals_map)
    aggregate_custom_by_po_item(data, cust_map, decimals_map=decimals_map)
    
    std_list = format_aggregation_as_list(std_map)
    cust_list = format_aggregation_as_list(cust_map)
    pallets_list = aggregate_per_po_with_pallets(data, decimals_map=decimals_map)
    
    daf_list = perform_DAF_compounding(
        cust_list,
        chunk_size=chunk_size,
        item_separator=item_separator,
        line_separator=line_separator,
        decimals_map=decimals_map
    )
    
    item_merged_map = {}
    if cust_list:
        aggregate_by_item_merge_po(cust_list, item_merged_map, decimals_map=decimals_map)
        item_merged_list = format_aggregation_as_list(item_merged_map)
    else:
        item_merged_list = []
        
    return {
        "standard": std_list,
        "custom": cust_list,
        "manifest_by_pallet_per_po": pallets_list,
        "daf": daf_list,
        "item_merged": item_merged_list,
    }


run_aggregation_pipeline = build_single_table

__all__ = [
    "Aggregator",
    "int_sum_reducer",
    "decimal_sum_reducer",
    "first_value_reducer",
    "first_non_empty_reducer",
    "calculate_footer_totals",
    "format_aggregation_as_list",
    "quantize_record",
    "finalize_standard_aggregation",
    "finalize_item_merged_aggregation",
    "aggregate_standard_by_po_item_price",
    "aggregate_custom_by_po_item",
    "aggregate_by_item_merge_po",
    "calculate_leather_summary",
    "aggregate_per_po_with_pallets",
    "perform_DAF_compounding",
    "DAFCompoundingResult",
    "FinalDAFResultType",
    "build_single_table",
    "coerce_input_row",
    "execute_strategy",
    "run_aggregation_pipeline",
    "INVOICE_LINE_ITEM_COLUMNS",
    "INVOICE_FOOTER_COLUMNS",
]
