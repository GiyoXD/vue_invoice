import logging
from typing import Any, Dict, List

logger = logging.getLogger(__name__)

def apply_vertical_merges(data_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Computes vertical cell merges for pallet numbers and descriptions.
    Appends a 'rowspans' metadata dict to each row (e.g. row['rowspans']['col_desc'] = 3).
    """
    if not data_rows:
        return data_rows

    # Determine which columns are allowed to merge
    merge_columns = ["col_pallet_no"]
    
    # Self-contained check: only merge descriptions if they are uniform across all rows
    descriptions = {str(r.get("col_desc", "")).strip().lower() for r in data_rows if r.get("col_desc")}
    if len(descriptions) <= 1:
        merge_columns.append("col_desc")
    else:
        logger.info(f"[MergeTransformer] Skipping col_desc merge: mixed descriptions found {descriptions}")

    # Track active merge groups: col_id -> {"start_idx": int, "value": Any}
    active_groups = {}

    for i, row in enumerate(data_rows):
        row["rowspans"] = {}
        for col_id in merge_columns:
            val = row.get(col_id)
            
            # Skip merging if value is an integer (pallet numbers that are digits shouldn't auto-merge)
            is_int = False
            if val is not None:
                try:
                    int(val)
                    is_int = True
                except (ValueError, TypeError):
                    pass

            if not is_int and col_id in active_groups and active_groups[col_id]["value"] == val and val is not None:
                row["rowspans"][col_id] = 0  # Swallowed by merge
                start_idx = active_groups[col_id]["start_idx"]
                data_rows[start_idx]["rowspans"][col_id] += 1
            else:
                # Start a new merge group
                active_groups[col_id] = {"start_idx": i, "value": val}
                row["rowspans"][col_id] = 1

    return data_rows
