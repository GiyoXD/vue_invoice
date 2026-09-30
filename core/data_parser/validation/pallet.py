import logging
import decimal
import re
from typing import List, Dict, Any, Optional
from .exceptions import DataValidationError
from .helpers import _is_valid_cell_value

decimal.getcontext().prec = 28


def _normalize_pallet_id(raw_id: Any) -> str:
    """Normalize pallet ID to stripped uppercase string."""
    if raw_id is None:
        return ""
    return str(raw_id).strip().upper()


def _parse_pallet_count(raw_count: Any) -> int:
    """
    Parses pallet count value into binary marker (1 or 0).
    Supports ints, floats, strings, and ranges like '1-19' or '1~19'
    (where first number >= 1 gives count = 1) without throwing ValueError.
    """
    if raw_count is None:
        return 0
    try:
        if isinstance(raw_count, (int, float, decimal.Decimal)):
            return 1 if float(raw_count) >= 1 else 0

        val_str = str(raw_count).strip()
        if not val_str:
            return 0

        try:
            return 1 if float(val_str.replace(',', '')) >= 1 else 0
        except (ValueError, TypeError):
            pass

        match = re.match(r'^([+-]?\d+(?:\.\d+)?)\s*(?:[-~–—/]|to\b|of\b|至|到)', val_str, re.IGNORECASE)
        if match:
            return 1 if float(match.group(1)) >= 1 else 0

        match = re.match(r'^([+-]?\d+(?:\.\d+)?)\s*(?:plt|pallet|托|箱|件|ctn)?$', val_str, re.IGNORECASE)
        if match:
            return 1 if float(match.group(1)) >= 1 else 0
    except Exception:
        return 0

    return 0


def verify_pallet_integrity(
    table_data: List[Dict[str, Any]],
    table_id_str: Optional[str] = None,
    monitor: Optional[Any] = None
) -> None:
    """
    Validates pallet data integrity in three phases:
    
    Phase 1 — Pairing check (col_pallet_id vs col_qty_sf):
      If col_pallet_id column exists, every row with sqft data MUST have a pallet ID.
      Missing pallet ID on a data row = WARNING (operator forgot to fill it).
      If col_pallet_id column doesn't exist at all = warn and skip.
    
    Phase 2 — Correlation check (col_pallet_count vs col_pallet_id):
      Only runs if BOTH columns exist.
      - If col_pallet_id changed from last row, col_pallet_count MUST be 1.
      - If col_pallet_id is same as last row, col_pallet_count MUST be 0.
      - If col_pallet_count is 1, a valid non-empty col_pallet_id MUST be present.
      - A col_pallet_id cannot reappear after a different one (no gaps).

    Phase 3 — Pallet Metric Alignment Check (col_cbm, col_net, col_gross):
      Checks col_cbm, col_net, and col_gross alignment with pallet anchor rows.
      When count == 0 (sub-row) and the row has non-empty/non-zero value for any of
      ['col_cbm', 'col_net', 'col_gross'], logs a warning.
    """
    pallet_count_key = 'col_pallet_count'
    pallet_id_key = 'col_pallet_id'
    sqft_key = 'col_qty_sf'
    
    has_count = any(pallet_count_key in row for row in table_data)
    has_id = any(pallet_id_key in row for row in table_data)
    
    if not has_id and not has_count:
        return

    # --- Phase 1: Pairing check (pallet_id must exist for every row with sqft) ---
    if has_id:
        missing_id_rows = []
        for idx, row in enumerate(table_data):
            row_num = row.get('_row_num', idx + 1)
            sqft_val = row.get(sqft_key)
            raw_id = row.get(pallet_id_key)
            pallet_id = _normalize_pallet_id(raw_id)
            
            # Row has sqft data but no pallet ID → operator forgot to fill
            has_sqft = False
            if sqft_val is not None and str(sqft_val).strip() != "":
                try:
                    dec_val = sqft_val if isinstance(sqft_val, decimal.Decimal) else decimal.Decimal(str(sqft_val).replace(',', '').strip())
                    if dec_val != 0:
                        has_sqft = True
                except (decimal.InvalidOperation, ValueError, TypeError):
                    has_sqft = True
            if has_sqft and not pallet_id:
                missing_id_rows.append(row_num)
        
        if missing_id_rows:
            row_list = ', '.join(str(r) for r in missing_id_rows[:10])
            suffix = f" (and {len(missing_id_rows) - 10} more)" if len(missing_id_rows) > 10 else ""
            msg = (
                f"[{table_id_str or 'Table'}] [Pallet Pairing] {len(missing_id_rows)} row(s) have sqft data but missing Pallet ID: "
                f"rows {row_list}{suffix}. Operator may have forgotten to fill these."
            )
            logging.warning(msg)
            if monitor:
                monitor.log_warning(msg)
    else:
        msg = f"[Pallet Pairing] col_pallet_id column not found in {table_id_str or 'data'}. Skipping pallet ID validation."
        logging.warning(msg)
        if monitor:
            monitor.log_warning(msg)

    # --- Phase 2: Correlation check (count vs id) — only if both columns exist ---
    if has_id and has_count:
        last_pallet_id = None
        seen_pallet_ids = set()

        for idx, row in enumerate(table_data):
            row_num = row.get('_row_num', idx + 1)
            raw_count = row.get(pallet_count_key, 0)
            count = _parse_pallet_count(raw_count)
            raw_id = row.get(pallet_id_key)
            pallet_id = _normalize_pallet_id(raw_id)
            
            # If both are empty, we might be on a non-pallet/footer/ignored row.
            # We skip validation without resetting last_pallet_id so spacer rows do not trigger false errors.
            if not pallet_id and count == 0:
                continue

            # Rule C (Presence): If count is 1, pallet_id must be present
            if count == 1 and not pallet_id:
                raise DataValidationError(
                    f"Row {row_num}: Pallet ID missing (count = 1)"
                )

            # Rules A & B: Transitions
            if pallet_id != last_pallet_id:
                # Count MUST be 1 (Rule A)
                if count != 1:
                    raise DataValidationError(
                        f"Row {row_num}: Pallet ID changed to '{pallet_id}' but count is {raw_count}"
                    )
                
                # Rule D: Contiguity (No recurrence after a gap)
                if pallet_id in seen_pallet_ids:
                    raise DataValidationError(
                        f"Row {row_num}: Pallet ID '{pallet_id}' reappeared after gap"
                    )
                seen_pallet_ids.add(pallet_id)
            else:
                # Continuation of the same pallet ID
                # Count MUST be 0 (Rule B)
                if count != 0:
                    raise DataValidationError(
                        f"Row {row_num}: Pallet count is 1 but Pallet ID did not change ('{pallet_id}')"
                    )

            last_pallet_id = pallet_id

    # --- Phase 3: Pallet Metric Alignment Check (col_cbm, col_net, col_gross on anchor vs sub-rows) ---
    if has_count:
        metric_cols = ['col_cbm', 'col_net', 'col_gross']
        col_name_map = {'col_cbm': 'CBM', 'col_net': 'Net Weight', 'col_gross': 'Gross Weight'}
        for idx, row in enumerate(table_data):
            row_num = row.get('_row_num', idx + 1)
            raw_count = row.get(pallet_count_key, 0)
            count = _parse_pallet_count(raw_count)

            raw_id = row.get(pallet_id_key)
            pallet_id = _normalize_pallet_id(raw_id)
            pallet_label = f" (Pallet '{pallet_id}')" if pallet_id else ""

            if count == 0 and (not has_id or pallet_id):
                for col in metric_cols:
                    val = row.get(col)
                    if _is_valid_cell_value(val, is_numeric=True):
                        metric_name = col_name_map.get(col, col)
                        msg = (
                            f"[{table_id_str or 'Table'}] Row {row_num}{pallet_label}: "
                            f"{metric_name} ({val}) is on a sub-row (Pallet=0). Should be on main pallet row (Pallet=1)."
                        )
                        logging.warning(msg)
                        if monitor:
                            monitor.log_warning(msg)
