
import logging
import decimal
from typing import List, Dict, Any, Optional
from .util.converters import DataConverter

# Set precision for Decimal calculations (consistent with data_processor)
decimal.getcontext().prec = 28

class DataValidationError(Exception):
    """User-facing validation error for missing required data.
    
    This exception is caught separately in the API layer to return
    a clean, human-readable error message without traceback noise.
    """
    pass


def validate_table_data_presence(
    current_table_data: List[Dict[str, Any]], 
    table_id_str: str, 
    column_mapping: Dict[str, str], 
    monitor: Optional[Any] = None
):
    """
    Validates that the first row of the table contains valid values for all required columns.
    If a required column is missing its value on the first row, throw DataValidationError.
    """
    # Columns that MUST always be present in every valid table
    ALWAYS_REQUIRED = [
        'col_po', 'col_item', 'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm'
    ]
    
    # Pricing columns — only required if the scanner actually found them.
    # Shipping lists (net-weight mode) won't have these; they get injected later.
    PRICING_COLUMNS = [
        'col_amount', 'col_unit_price', 'col_qty_sf'
    ]
    
    if not current_table_data:
        logging.warning(f"Validation: {table_id_str} has no data rows to validate.")
        return

    first_row = current_table_data[0]
    missing_data_cols = []
    
    # Build the actual required list: always-required + pricing cols IF mapped
    required_cols = list(ALWAYS_REQUIRED)
    for pc in PRICING_COLUMNS:
        if pc in column_mapping:
            required_cols.append(pc)
    
    for col_name in required_cols:
        # 1. Check mapping first (did we even find the header?)
        if col_name not in column_mapping:
            missing_data_cols.append(f"{col_name} (Missing Header)")
            continue
        
        # 2. Check ONLY the first row for a valid value
        val = first_row.get(col_name)
        has_val = False
        if val is not None:
             if isinstance(val, (int, float, decimal.Decimal)):
                 if val != 0:
                     has_val = True
             elif str(val).strip():
                 has_val = True
        
        if not has_val:
            missing_data_cols.append(col_name)

    if missing_data_cols:
        err_msg = (
            f"Data Validation Error: {table_id_str} is missing mandatory data in the first row for: "
            f"[{', '.join(missing_data_cols)}]. "
            "Please ensure the first row of every table in your Excel is fully populated."
        )
        if monitor:
            monitor.log_process_item(f"{table_id_str} First-Row Validation", status="error", error=err_msg)
        raise DataValidationError(err_msg)

def validate_weight_integrity(data_rows: List[Dict[str, Any]], monitor: Optional[Any] = None):
    """
    Strict validation to ensure Gross Weight is always strictly bigger than Net Weight.
    Also verifies that the Tare weight (Gross - Net) is consistent across rows.
    """
    prefix = "[validate_weight_integrity]"
    
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
    ref_row_info: str = ""

    for i, row in enumerate(data_rows):
        # Get raw values
        net_raw = row.get(net_key)
        gross_raw = row.get(gross_key)
        
        # --- NEW STRICT VALIDATION: Unpaired weights are forbidden ---
        # If one exists but the other doesn't, it's a data entry error
        if (net_raw is not None and gross_raw is None) or (net_raw is None and gross_raw is not None):
            po_val = row.get(po_key, "Unknown PO")
            item_val = row.get(item_key, "Unknown Item")
            missing_col = "Gross Weight" if gross_raw is None else "Net Weight"
            present_col = "Net Weight" if net_raw is not None else "Gross Weight"
            
            error_msg = (
                f"Weight Integrity Error: Partial weight found at row for PO [{po_val}] / Item [{item_val}]. "
                f"{present_col} has a value, but {missing_col} is missing. "
                "Weights must always be provided as a Net/Gross pair."
            )
            logging.error(f"{prefix} {error_msg}")
            raise DataValidationError(error_msg)

        # Skip if both are missing (legitimate filler or spacer row)
        if net_raw is None and gross_raw is None:
            continue
            
        try:
            # Handle Decimals, floats, or strings
            net_val = net_raw if isinstance(net_raw, decimal.Decimal) else DataConverter.convert_to_decimal(net_raw)
            gross_val = gross_raw if isinstance(gross_raw, decimal.Decimal) else DataConverter.convert_to_decimal(gross_raw)
            
            if net_val is None or gross_val is None:
                # This handles cases where values might be empty strings after stripping (if convert_to_decimal returns None)
                # We treat this as missing/filler unless it's an unpaired entry
                continue

            # Skip header/footer rows where BOTH are 0
            if net_val == 0 and gross_val == 0:
                continue

            # Validation 1: Strict Positivity Constraint (Gross > Net)
            if gross_val <= net_val:
                po_val = row.get(po_key, "Unknown PO")
                item_val = row.get(item_key, "Unknown Item")
                
                error_msg = (
                    f"Weight Validation Error: At row for PO [{po_val}] / Item [{item_val}], "
                    f"Gross Weight ({gross_val}) is not strictly greater than Net Weight ({net_val}). "
                    "In shipping, Gross Weight MUST always be bigger than Net Weight. "
                    "Please fix your source Excel and try again."
                )
                
                logging.error(f"{prefix} {error_msg}")
                raise DataValidationError(error_msg)

            # Validation 2: Tare Weight Consistency
            current_tare = gross_val - net_val
            
            po_val = row.get(po_key, "Unknown PO")
            item_val = row.get(item_key, "Unknown Item")
            row_id_str = f"PO [{po_val}] / Item [{item_val}]"

            if reference_tare is None:
                reference_tare = current_tare
                ref_row_info = row_id_str
                logging.info(f"{prefix} Established reference tare weight: {reference_tare} from {ref_row_info}")
            else:
                if current_tare != reference_tare:
                    expected_gross = net_val + reference_tare
                    error_msg = (
                        f"Weight Integrity Error: `Net + Pallet Weight` does not equal `Gross Weight` at {row_id_str}. "
                        f"Based on the first row ({ref_row_info}), the Pallet Weight (Tare) is **{reference_tare}**. "
                        f"Expected Gross Weight for this row is {net_val} + {reference_tare} = **{expected_gross}**, "
                        f"but found **{gross_val}**. "
                        "Please ensure all pallets in your table have identical tare weights."
                    )
                    logging.error(f"{prefix} {error_msg}")
                    raise DataValidationError(error_msg)

        except (decimal.InvalidOperation, ValueError, TypeError):
            continue

def validate_data(
    data_rows: List[Dict[str, Any]], 
    table_id_str: str, 
    column_mapping: Dict[str, str], 
    monitor: Optional[Any] = None,
    phase: str = 'presence'
):
    """
    Unified entry point for all table-level data validation.
    
    Args:
        data_rows: The list of row dictionaries to validate.
        table_id_str: Human readable ID (e.g. "Table 1") for error messages.
        column_mapping: The canonical-to-Excel-letter mapping.
        monitor: Optional PipelineMonitor for logging errors.
        phase: 'presence' (check for at least one value per required col) 
               or 'integrity' (check weight consistency and tare matching).
    """
    if phase == 'presence':
        validate_table_data_presence(data_rows, table_id_str, column_mapping, monitor=monitor)
    elif phase == 'integrity':
        # Weight integrity is a data-level check, it doesn't need column mapping
        # but we use table_id_str in logs via monitor if provided (future enhancement)
        validate_weight_integrity(data_rows, monitor=monitor)
    else:
        raise ValueError(f"Unknown validation phase requested: {phase}. Supported: 'presence', 'integrity'.")


logging.info("[validation] Module loaded with consolidated validation routines.")
