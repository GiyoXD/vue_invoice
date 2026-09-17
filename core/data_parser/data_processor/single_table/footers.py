from typing import List, Dict, Any, Optional
from .schema import INVOICE_FOOTER_COLUMNS, COLUMN_REDUCERS
from .aggregations import quantize_record
from .engine import coerce_input_row

__all__ = ["calculate_footer_totals"]


def calculate_footer_totals(
    processed_data: List[Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None
) -> Dict[str, Any]:
    """Calculates totals for footer fields based on processed data."""
    processed_data = [coerce_input_row(r) for r in processed_data if r]
    totals = {
        col: COLUMN_REDUCERS[col]([r.get(col) for r in processed_data])
        for col in INVOICE_FOOTER_COLUMNS
    }
    p_cnt = totals['col_pallet_count']
    totals['multiple'] = "S" if p_cnt > 1 else ""

    quantize_record(totals, decimals_map)
    return totals
