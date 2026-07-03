import logging
import decimal
import re
from typing import Dict, List, Any, Optional, Tuple
from ..util.converters import DataConverter

# Define precision specifically for CBM results
CBM_DECIMAL_PLACES = decimal.Decimal('0.0001')

def _calculate_single_cbm(cbm_value: Any, row_index: int) -> Optional[decimal.Decimal]:
    """
    Parses a CBM string (e.g., "L*W*H" or "LxWxH") and calculates the volume.
    """
    prefix = "[_calculate_single_cbm]"
    log_context = f"for CBM at row index {row_index}"

    if cbm_value is None:
        logging.debug(f"{prefix} Input CBM value is None. {log_context}")
        return None

    if isinstance(cbm_value, (int, float, decimal.Decimal)):
        logging.debug(f"{prefix} Input CBM is already numeric: {cbm_value}. {log_context}")
        calculated = DataConverter.convert_to_decimal(cbm_value, log_context)
        if calculated is not None:
             result = calculated.quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
             logging.debug(f"{prefix} Quantized pre-numeric CBM to {result}. {log_context}")
             return result
        else:
             logging.warning(f"{prefix} Failed to convert pre-numeric CBM value {cbm_value} to Decimal. {log_context}")
             return None

    if not isinstance(cbm_value, str):
        logging.warning(f"{prefix} Unexpected type '{type(cbm_value).__name__}' for CBM value '{cbm_value}'. Cannot calculate. {log_context}")
        return None

    cbm_str = cbm_value.strip()
    if not cbm_str:
        logging.debug(f"{prefix} Input CBM string is empty after strip. {log_context}")
        return None

    logging.debug(f"{prefix} Attempting to parse CBM string: '{cbm_str}'. {log_context}")

    parts = cbm_str.split('*')
    separator_used = "'*'"

    if len(parts) != 3:
        if '*' not in cbm_str and ('x' in cbm_str.lower()):
             parts = re.split(r'[xX]', cbm_str)
             separator_used = "'x' or 'X'"
             logging.debug(f"{prefix} Split by '*' failed, trying split by {separator_used}. Parts: {parts}. {log_context}")

    if len(parts) != 3:
        logging.warning(f"{prefix} Invalid CBM format: '{cbm_str}'. Expected 3 parts separated by '*' or 'x'. Found {len(parts)} parts: {parts}. {log_context}")
        return None

    try:
        dims = []
        valid_dims = True
        for i, part in enumerate(parts):
             dim = DataConverter.convert_to_decimal(part, f"{log_context}, part {i+1} ('{part}')")
             if dim is None:
                 logging.warning(f"{prefix} Failed to convert dimension part {i+1} ('{part}') to Decimal. {log_context}")
                 valid_dims = False
             dims.append(dim)

        if not valid_dims:
            logging.warning(f"{prefix} Failed to convert one or more dimensions for CBM string '{cbm_str}'. Cannot calculate volume. {log_context}")
            return None

        dim1, dim2, dim3 = dims
        volume = (dim1 * dim2 * dim3).quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
        logging.debug(f"{prefix} Calculated CBM volume: {volume} from '{cbm_str}' (Dims: {dims}). {log_context}")
        return volume

    except Exception as e:
        logging.error(f"{prefix} Unexpected error calculating CBM from '{cbm_str}': {e}. {log_context}", exc_info=True)
        return None


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

    if not any(cbm_key in row for row in raw_data):
        logging.debug(f"{prefix} No '{cbm_key}' column found in this table's extracted data. Skipping CBM calculation.")
        return raw_data

    logging.info(f"{prefix} Processing '{cbm_key}' column for volume calculations (Rows: {len(raw_data)})...")

    for i, row in enumerate(raw_data):
        if cbm_key in row:
            value = row[cbm_key]
            
            if isinstance(value, str):
                row['col_cbm_raw'] = value.strip()
            
            calculated_value = _calculate_single_cbm(value, i)
            row[cbm_key] = calculated_value

    logging.info(f"{prefix} Finished processing '{cbm_key}' column. Rows updated with calculated values (Decimals or Nones).")
    return raw_data
