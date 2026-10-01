import logging
import decimal
from typing import List, Dict, Any, Optional
from ..util.converters import DataConverter
from .exceptions import DataValidationError

decimal.getcontext().prec = 28


def validate_weight_integrity(
    data_rows: List[Dict[str, Any]], 
    table_id_str: Optional[str] = None, 
    monitor: Optional[Any] = None, 
    ignore_tare_warning: bool = False
):
    """
    Strict validation to ensure Gross Weight is always strictly bigger than Net Weight.
    Also verifies that the Tare weight (Gross - Net) is consistent across rows.
    """
    table_context = f" in {table_id_str}" if table_id_str else ""
    prefix = f"[validate_weight_integrity{table_context}]"
    
    # Column keys
    net_key = 'col_net'
    gross_key = 'col_gross'
    po_key = 'col_po'
    item_key = 'col_item'

    has_net = any(net_key in row for row in data_rows)
    has_gross = any(gross_key in row for row in data_rows)

    if not (has_net and has_gross):
        logging.debug(f"{prefix} Net or Gross column missing, skipping integrity check.")
        return

    reference_tare: Optional[decimal.Decimal] = None

    for idx, row in enumerate(data_rows):
        # Get raw values
        net_raw = row.get(net_key)
        gross_raw = row.get(gross_key)

        # Handle Decimals, floats, or strings
        net_val = net_raw if isinstance(net_raw, decimal.Decimal) else DataConverter.convert_to_decimal(net_raw)
        gross_val = gross_raw if isinstance(gross_raw, decimal.Decimal) else DataConverter.convert_to_decimal(gross_raw)
        
        # Skip if both are missing (legitimate filler or spacer row)
        if net_val is None and gross_val is None:
            continue

        # --- NEW STRICT VALIDATION: Unpaired weights are forbidden ---
        # If one exists but the other doesn't, it's a data entry error
        if net_val is None or gross_val is None:
            row_num = row.get('_row_num', idx + 1)
            if gross_val is None:
                error_msg = f"Row {row_num}: Missing Gross Weight (Net = {net_val:.2f})"
            else:
                error_msg = f"Row {row_num}: Missing Net Weight (Gross = {gross_val:.2f})"
            logging.error(f"{prefix} {error_msg}")
            raise DataValidationError(error_msg)

        # Skip header/footer rows where BOTH are 0
        if net_val == 0 and gross_val == 0:
            continue

        # Validation 1: Strict Positivity Constraint (Gross > Net)
        if gross_val <= net_val:
            row_num = row.get('_row_num', idx + 1)
            error_msg = f"Row {row_num}: Gross <= Net (Net = {net_val:.2f}, Gross = {gross_val:.2f})"
            logging.error(f"{prefix} {error_msg}")
            raise DataValidationError(error_msg)

        # Validation 2: Tare Weight Consistency
        current_tare = gross_val - net_val

        if reference_tare is None:
            reference_tare = current_tare
            po_val = row.get(po_key, "Unknown PO")
            item_val = row.get(item_key, "Unknown Item")
            row_id_str = f"PO [{po_val}] / Item [{item_val}]"
            logging.info(f"{prefix} Established reference tare weight: {reference_tare} from {row_id_str}")
        else:
            if current_tare != reference_tare:
                expected_gross = net_val + reference_tare
                row_num = row.get('_row_num', idx + 1)
                error_msg = f"Row {row_num}: Tare mismatch (expected Gross = {expected_gross:.2f}, but found {gross_val:.2f})"
                if not ignore_tare_warning:
                    logging.error(f"{prefix} {error_msg}")
                    raise DataValidationError(error_msg)
                else:
                    logging.warning(f"{prefix} Ignored: {error_msg}")
                    if monitor:
                        monitor.log_warning(f"Ignored: {error_msg}")
