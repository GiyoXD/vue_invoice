# Facade for the data_processor package
from ..validation import DataValidationError, verify_pallet_integrity
from .distribution import ProcessingError, distribute_values, inject_net_weight_pricing
from .cbm import process_cbm_column
from .pallet import normalize_pallet_count, format_pallet_counts_to_xy
from .normalization import normalize_table_types
from .single_table import (
    Aggregator,
    aggregate_standard_by_po_item_price,
    aggregate_custom_by_po_item,
    aggregate_by_item_merge_po,
    aggregate_per_po_with_pallets,
    calculate_leather_summary,
    calculate_footer_totals,
    format_aggregation_as_list,
    perform_DAF_compounding,
    execute_strategy,
    run_aggregation_pipeline,
    build_single_table,
)
from .footer import calculate_all_footers, format_footer_payload

__all__ = [
    "DataValidationError",
    "verify_pallet_integrity",
    "ProcessingError",
    "distribute_values",
    "inject_net_weight_pricing",
    "process_cbm_column",
    "normalize_pallet_count",
    "format_pallet_counts_to_xy",
    "normalize_table_types",
    "Aggregator",
    "aggregate_standard_by_po_item_price",
    "aggregate_custom_by_po_item",
    "aggregate_by_item_merge_po",
    "aggregate_per_po_with_pallets",
    "calculate_leather_summary",
    "calculate_footer_totals",
    "format_aggregation_as_list",
    "perform_DAF_compounding",
    "execute_strategy",
    "run_aggregation_pipeline",
    "build_single_table",
    "calculate_all_footers",
    "format_footer_payload",
]


