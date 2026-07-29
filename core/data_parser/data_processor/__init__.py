# Facade for the data_processor package
from ..validation import DataValidationError, verify_pallet_integrity
from .distribution import ProcessingError, distribute_values, normalize_by_pallet_anchor, inject_net_weight_pricing
from .cbm import process_cbm_column
from .pallet import normalize_pallet_count, format_pallet_counts_to_xy
from .normalization import normalize_table_types
from .aggregation import (
    Aggregator,
    aggregate_standard_by_po_item_price,
    aggregate_custom_by_po_item,
    aggregate_per_po_with_pallets,
    calculate_leather_summary,
    calculate_weight_summary,
    calculate_pallet_summary,
    calculate_footer_totals,
    format_aggregation_as_list,
    perform_DAF_compounding,
)
from .footer import calculate_all_footers

