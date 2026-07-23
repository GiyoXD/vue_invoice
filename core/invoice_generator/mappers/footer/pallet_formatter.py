import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


def format_pallet_counts(
    data_rows: List[Dict[str, Any]],
    num_data_rows: int,
    pallet_col_id: Optional[str]
) -> None:
    """
    Carries non-empty values forward across rows for vertical cell merging.
    Does not manufacture string values. Modifies data_rows in-place.
    """
    if num_data_rows <= 0 or not pallet_col_id:
        return

    carry_value = None
    for row in data_rows[:num_data_rows]:
        raw_val = row.get(pallet_col_id)
        if raw_val not in (None, "", 0):
            carry_value = raw_val
        elif carry_value is not None:
            row[pallet_col_id] = carry_value
