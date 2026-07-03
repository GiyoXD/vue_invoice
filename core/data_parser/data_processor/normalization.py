import decimal
import logging
from typing import List, Dict, Any
from ..util.converters import DataConverter

def normalize_table_types(table_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Normalizes raw table row values into clean Python types at the entry point of the pipeline.
    - Decimals: col_qty_sf, col_amount, col_net, col_gross, col_cbm, col_unit_price
    - Integers: col_qty_pcs
    This avoids redundant type checking and casting downstream.
    """
    decimal_cols = {
        'col_qty_sf', 'col_amount', 'col_net', 'col_gross', 'col_cbm', 'col_unit_price'
    }
    integer_cols = {
        'col_qty_pcs'
    }

    for row in table_data:
        # Convert decimal columns
        for col in decimal_cols:
            if col in row:
                val = row[col]
                if val is not None:
                    dec_val = DataConverter.convert_to_decimal(val, col)
                    row[col] = dec_val

        # Convert integer columns
        for col in integer_cols:
            if col in row:
                val = row[col]
                if val is not None:
                    try:
                        # Strip commas and handle float strings (like "137.0")
                        row[col] = int(float(str(val).replace(',', '').strip()))
                    except (ValueError, TypeError):
                        row[col] = None

    return table_data
