import logging
import decimal
import re
from typing import Dict, List, Any, Optional, Tuple
from ..util.converters import DataConverter

# Define precision specifically for CBM results
CBM_DECIMAL_PLACES = decimal.Decimal('0.01')

def _quantize_cbm(d: Optional[decimal.Decimal]) -> Optional[decimal.Decimal]:
    if d is None or d.is_nan() or d.is_infinite():
        return None
    try:
        return d.quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
    except (decimal.InvalidOperation, decimal.OverflowError):
        return None


def _calculate_single_cbm(cbm_value: Any, row_index: int) -> Optional[decimal.Decimal]:
    """
    Parses a CBM string (e.g., "L*W*H" or "LxWxH") and calculates the volume.
    """
    prefix = "[_calculate_single_cbm]"
    log_context = f"for CBM at row index {row_index}"

    if cbm_value is None:
        return None

    if isinstance(cbm_value, (int, float, decimal.Decimal)):
        calculated = DataConverter.convert_to_decimal(cbm_value, log_context)
        if calculated is not None:
            return _quantize_cbm(calculated)
        else:
            logging.warning(f"{prefix} Failed to convert pre-numeric CBM value {cbm_value} to Decimal. {log_context}")
            return None

    cbm_str = str(cbm_value).strip().lstrip('=').strip()
    if not cbm_str:
        return None
    parts = [p.strip() for p in re.split(r'[*xX]', cbm_str)]
    if len(parts) == 3 and all(parts):
        dims = []
        valid_dims = True
        for i, part in enumerate(parts):
            dim = DataConverter.convert_to_decimal(part, f"{log_context}, part {i+1} ('{part}')")
            if dim is None or dim.is_nan() or dim.is_infinite():
                valid_dims = False
                break
            dims.append(dim)
        if valid_dims:
            try:
                prod = dims[0] * dims[1] * dims[2]
                return _quantize_cbm(prod)
            except (decimal.InvalidOperation, decimal.OverflowError):
                return None
    # If not 3 parts or dims invalid, attempt single numeric decimal conversion
    dec_val = DataConverter.convert_to_decimal(cbm_str, log_context)
    if dec_val is not None:
        return _quantize_cbm(dec_val)
    return None


def _eval_cbm_formula(s: Any, row_idx: int = 0) -> Tuple[Optional[decimal.Decimal], Optional[str]]:
    if not s:
        return None, None
    s_clean = str(s).strip().lstrip('=').strip()
    parts = [p.strip() for p in re.split(r'[*xX]', s_clean)]
    if len(parts) == 3 and all(parts):
        val = _calculate_single_cbm(s_clean, row_idx)
        if val is not None:
            return val, s_clean
    return None, None


def process_cbm_column(raw_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Iterates through each row in raw_data, calculates numeric CBM values
    from strings (L*W*H or LxWxH format), and updates the row dict in place.
    """
    prefix = "[process_cbm_column]"
    cbm_key = 'col_cbm'

    if not raw_data:
        logging.debug(f"{prefix} Input data is empty. Skipping CBM calculation.")
        return raw_data

    has_cbm = any((row.get('col_cbm') is not None or row.get('col_cbm_raw') is not None) for row in raw_data)
    if not has_cbm:
        logging.debug(f"{prefix} No '{cbm_key}' column found in this table's extracted data. Skipping CBM calculation.")
        return raw_data

    for i, row in enumerate(raw_data):
        raw_formula = row.get('col_cbm_raw')
        cbm_val = row.get(cbm_key)

        # If cbm_val is formula string
        cbm_f_val, cbm_f_str = _eval_cbm_formula(cbm_val, i) if cbm_val is not None else (None, None)
        if cbm_f_val is not None:
            row[cbm_key] = cbm_f_val
            row['col_cbm_raw'] = cbm_f_str
            continue

        # If cbm_val is numeric
        dec_val = _calculate_single_cbm(cbm_val, i) if cbm_val is not None else None
        if dec_val is not None:
            row[cbm_key] = dec_val
            # KEEP row['col_cbm_raw'] if present! Only if row['col_cbm_raw'] is not set, leave it as is
            continue

        # cbm_val is None; check raw_formula
        if raw_formula is not None:
            rf_val, rf_str = _eval_cbm_formula(raw_formula, i)
            row[cbm_key] = rf_val if rf_val is not None else _calculate_single_cbm(raw_formula, i)
            row['col_cbm_raw'] = raw_formula
        else:
            row[cbm_key] = None

    calculated_count = sum(1 for row in raw_data if row.get(cbm_key) is not None)
    logging.debug(f"[CBM] Calculated volume for {calculated_count} rows")
    return raw_data
