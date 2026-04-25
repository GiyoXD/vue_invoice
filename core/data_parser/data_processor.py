
import logging
from typing import Dict, List, Any, Optional, Tuple
import decimal # Use Decimal for precise calculations
import re
import pprint
# Import config values (consider passing as arguments)
from .config import DISTRIBUTION_BASIS_COLUMN # Keep this

# Set precision for Decimal calculations
decimal.getcontext().prec = 28 # Default precision, adjust if needed
# Define precision specifically for CBM results (e.g., 4 decimal places)
CBM_DECIMAL_PLACES = decimal.Decimal('0.0001')
# Define default precision for other distributions (e.g., 4 decimal places)
DEFAULT_DIST_PRECISION = decimal.Decimal('0.0001')
from .validation import DataValidationError
from .util.converters import DataConverter
# Safely link the local legacy name to the new centralized utility
_convert_to_decimal = DataConverter.convert_to_decimal


class ProcessingError(Exception):
    """Custom exception for data processing errors."""
    pass


# _convert_to_decimal is now imported from .validation

# _calculate_single_cbm function remains unchanged...
def _calculate_single_cbm(cbm_value: Any, row_index: int) -> Optional[decimal.Decimal]:
    """
    Parses a CBM string (e.g., "L*W*H" or "LxWxH") and calculates the volume.

    Args:
        cbm_value: The value from the CBM cell (can be string, number, None).
        row_index: The 0-based index of the row for logging purposes.

    Returns:
        The calculated CBM as a Decimal, or None if parsing fails or input is invalid.
    """
    prefix = "[_calculate_single_cbm]"
    log_context = f"for CBM at row index {row_index}" # Use 0-based index internally

    if cbm_value is None:
        logging.debug(f"{prefix} Input CBM value is None. {log_context}")
        return None

    # If it's already a number, convert to Decimal and quantize
    if isinstance(cbm_value, (int, float, decimal.Decimal)):
        logging.debug(f"{prefix} Input CBM is already numeric: {cbm_value}. {log_context}")
        calculated = DataConverter.convert_to_decimal(cbm_value, log_context)
        if calculated is not None:
             result = calculated.quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
             logging.debug(f"{prefix} Quantized pre-numeric CBM to {result}. {log_context}")
             return result
        else:
             # Conversion should ideally not fail here, but handle it
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

    # Try splitting by '*' first
    parts = cbm_str.split('*')
    separator_used = "'*'"

    # If not 3 parts, try splitting by 'x' or 'X' (case-insensitive)
    if len(parts) != 3:
        if '*' not in cbm_str and ('x' in cbm_str.lower()):
             parts = re.split(r'[xX]', cbm_str) # Split by 'x' or 'X'
             separator_used = "'x' or 'X'"
             logging.debug(f"{prefix} Split by '*' failed, trying split by {separator_used}. Parts: {parts}. {log_context}")

    # Check if we have exactly 3 parts after trying separators
    if len(parts) != 3:
        logging.warning(f"{prefix} Invalid CBM format: '{cbm_str}'. Expected 3 parts separated by '*' or 'x'. Found {len(parts)} parts: {parts}. {log_context}")
        return None

    try:
        # Convert each part to Decimal
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


def inject_net_weight_pricing(tables: List[List[Dict[str, Any]]], global_unit_price: float) -> List[List[Dict[str, Any]]]:
    """
    For 'net' pricing mode: injects col_unit_price, col_amount, and col_qty_sf
    from col_net and the user-provided global unit price.
    
    This allows shipping lists that only have Net Weight (no sqft/price/amount)
    to flow through the standard aggregation and invoice generation pipeline.
    
    Args:
        tables: List of tables, each a list of row dicts (the 'multi_table' structure).
        global_unit_price: The user-provided unit price (USD per kg).
    
    Returns:
        The same tables with injected columns.
    """
    prefix = "[inject_net_weight_pricing]"
    price = decimal.Decimal(str(global_unit_price))
    injected_count = 0

    for table in tables:
        if not isinstance(table, list):
            continue
        for row in table:
            if not isinstance(row, dict):
                continue
            net_val = row.get('col_net')
            if net_val is not None:
                net_dec = _convert_to_decimal(net_val)
                if net_dec is not None and net_dec > 0:
                    row['col_unit_price'] = float(price)
                    row['col_amount'] = float((net_dec * price).quantize(
                        decimal.Decimal('0.01'), rounding=decimal.ROUND_HALF_UP
                    ))
                    # Use net weight as quantity basis if sqft is missing
                    if row.get('col_qty_sf') is None:
                        row['col_qty_sf'] = float(net_dec)
                    injected_count += 1


    logging.info(f"{prefix} Injected unit_price={price} into {injected_count} rows across {len(tables)} tables.")
    return tables


def inject_flat_net_weight_pricing(rows: List[Dict[str, Any]], global_unit_price: float) -> List[Dict[str, Any]]:
    """
    Applies net weight pricing to a flat list of dicts (like aggregated results).
    Re-calculates col_amount and col_qty_sf based on col_net sum.
    """
    price = decimal.Decimal(str(global_unit_price))
    for row in rows:
        if not isinstance(row, dict): continue
        net_val = row.get('col_net')
        if net_val is not None:
            net_dec = _convert_to_decimal(net_val)
            if net_dec is not None and net_dec > 0:
                row['col_unit_price'] = float(price)
                row['col_amount'] = float((net_dec * price).quantize(
                    decimal.Decimal('0.01'), rounding=decimal.ROUND_HALF_UP
                ))
                # For aggregated output from main.py, missing sf might be 0 or "0"
                if row.get('col_qty_sf') is None or row.get('col_qty_sf') in [0, "0", 0.0, "0.0"]:
                    row['col_qty_sf'] = float(net_dec)
    return rows


def normalize_pallet_count(table_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Normalizes pallet count values from 'x-y' format (e.g. '1-39', '2-39')
    into simple integers (1 or 0).
    
    Logic:
    - Parse 'x-y' to extract pallet number x.
    - First occurrence of a new x → 1 (new pallet)
    - Subsequent consecutive rows with same x → 0 (continuation row)
    - Values that are already plain integers are left as-is.
    
    Args:
        table_data: List of row dicts from a single extracted table.
    
    Returns:
        The same list with col_pallet_count normalized in-place.
    """
    prefix = "[normalize_pallet_count]"
    pallet_key = 'col_pallet_count'
    
    if not table_data:
        return table_data
    
    # Check if any value is in 'x-y' format
    sample_val = None
    for row in table_data:
        v = row.get(pallet_key)
        if v is not None and str(v).strip():
            sample_val = str(v).strip()
            break
    
    if sample_val is None:
        return table_data
    
    # Only activate if the format is 'x-y' (contains a dash with digits on both sides)
    import re
    if not re.match(r'^\d+-\d+$', sample_val):
        logging.debug(f"{prefix} Pallet values are not in 'x-y' format (sample: '{sample_val}'). Skipping.")
        return table_data
    
    logging.info(f"{prefix} Detected 'x-y' pallet format (sample: '{sample_val}'). Normalizing...")
    
    last_pallet_num = None
    normalized_count = 0
    
    for row in table_data:
        raw_val = row.get(pallet_key)
        if raw_val is None or not str(raw_val).strip():
            row[pallet_key] = 0
            continue
        
        raw_str = str(raw_val).strip()
        match = re.match(r'^(\d+)-(\d+)$', raw_str)
        
        if match:
            pallet_num = int(match.group(1))
            
            if pallet_num != last_pallet_num:
                # New pallet
                row[pallet_key] = 1
                last_pallet_num = pallet_num
                normalized_count += 1
            else:
                # Continuation row (same pallet)
                row[pallet_key] = 0
        else:
            # Not x-y format, try plain int conversion
            try:
                row[pallet_key] = int(float(raw_str))
            except (ValueError, TypeError):
                row[pallet_key] = 0
    
    logging.info(f"{prefix} Normalized {len(table_data)} rows → {normalized_count} unique pallets detected.")
    return table_data

# process_cbm_column function remains unchanged...
def process_cbm_column(raw_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Iterates through each row in raw_data, calculates numeric CBM values
    from strings (L*W*H or LxWxH format), and updates the row dict in place.
    """
    prefix = "[process_cbm_column]"
    cbm_key = 'col_cbm' # Canonical name

    if not raw_data:
        logging.debug(f"{prefix} Input data is empty. Skipping CBM calculation.")
        return raw_data

    # Check if cbm_key exists in at least one row
    if not any(cbm_key in row for row in raw_data):
        logging.debug(f"{prefix} No '{cbm_key}' column found in this table's extracted data. Skipping CBM calculation.")
        return raw_data

    logging.info(f"{prefix} Processing '{cbm_key}' column for volume calculations (Rows: {len(raw_data)})...")

    # Process each row in the list
    for i, row in enumerate(raw_data):
        if cbm_key in row:
            value = row[cbm_key]
            
            # Save the raw string before calculating
            if isinstance(value, str):
                row['col_cbm_raw'] = value.strip()
            
            calculated_value = _calculate_single_cbm(value, i) # Calculate volume using the helper
            row[cbm_key] = calculated_value # Replace string with Decimal or None

    logging.info(f"{prefix} Finished processing '{cbm_key}' column. Rows updated with calculated values (Decimals or Nones).")
    return raw_data

from core.utils.pipeline_monitor import PipelineMonitor

def normalize_by_pallet_anchor(
    raw_data: List[Dict[str, Any]],
    columns_to_normalize: List[str],
    basis_column: str = 'col_qty_pcs',
    pallet_column: str = 'col_pallet_count',
    monitor: Optional[PipelineMonitor] = None
) -> List[Dict[str, Any]]:
    """
    Uses col_pallet_count as anchor to fix misplaced distributable values.

    Rule: net/gross/cbm are ONLY valid on the same row as pallet_count.
    If they appear on a row without pallet_count, they are misplaced and
    must be pulled up to the nearest pallet anchor row above.

    After pull-up:
    - If the filler row still has pcs (basis), keep it but clear its distributable values.
    - If the filler row has no pcs, remove it entirely.

    Args:
        raw_data: List of row dicts (post-CBM processing).
        columns_to_normalize: Columns to check/pull up (e.g. ['col_net', 'col_gross', 'col_cbm']).
        basis_column: The basis column name (e.g. 'col_qty_pcs') used to decide if a filler row has data.
        pallet_column: The pallet count column name.

    Returns:
        Cleaned list of row dicts with values consolidated onto anchor rows.
    """
    prefix = "[normalize_by_pallet_anchor]"

    if not raw_data:
        return []

    # --- Resolve canonical column names ---
    # We strictly enforce the 'col_' prefix for all canonical column names.
    # Magic auto-prefixing has been removed to ensure configuration clarity.
    for col in columns_to_normalize:
        if not col.startswith('col_'):
            raise ValueError(f"Normalization Error: Column name '{col}' must start with 'col_'.")
    
    if not basis_column.startswith('col_'):
        raise ValueError(f"Normalization Error: Basis column '{basis_column}' must start with 'col_'.")

    resolved_cols = list(columns_to_normalize)
    resolved_basis = basis_column

    # --- Check if pallet_column exists in data at all ---
    has_pallet_col = any(pallet_column in row for row in raw_data)
    if not has_pallet_col:
        logging.info(f"{prefix} No '{pallet_column}' found in data. Skipping normalization.")
        return raw_data

    # --- Helper: check if a row is a pallet anchor ---
    def _is_pallet_row(row: Dict[str, Any]) -> bool:
        """Returns True if this row has a pallet_count value >= 1."""
        val = row.get(pallet_column)
        if val is None:
            return False
        try:
            return decimal.Decimal(str(val)) >= 1
        except (decimal.InvalidOperation, ValueError, TypeError):
            return False

    def _has_distributable_value(row: Dict[str, Any]) -> bool:
        """Returns True if any of the distributable columns have a non-empty, non-zero value."""
        for col in resolved_cols:
            val = row.get(col)
            if val is not None:
                try:
                    if decimal.Decimal(str(val)) != 0:
                        return True
                except (decimal.InvalidOperation, ValueError, TypeError):
                    # Non-numeric but non-None (e.g. a string) — treat as having a value
                    if str(val).strip():
                        return True
        return False

    def _has_basis_value(row: Dict[str, Any]) -> bool:
        """Returns True if the row has a valid, non-zero basis (pcs) value."""
        val = row.get(resolved_basis)
        if val is None:
            return False
        try:
            return decimal.Decimal(str(val)) > 0
        except (decimal.InvalidOperation, ValueError, TypeError):
            return False

    # --- Build pallet groups ---
    # Each group: { 'anchor_idx': int, 'member_indices': [int, ...] }
    groups = []
    current_group = None

    for i, row in enumerate(raw_data):
        if _is_pallet_row(row):
            current_group = {'anchor_idx': i, 'member_indices': [i]}
            groups.append(current_group)
        elif current_group is not None:
            current_group['member_indices'].append(i)
        else:
            # Rows before the first pallet anchor — no group to attach to, leave as-is
            pass

    if not groups:
        logging.info(f"{prefix} No pallet anchor rows found. Skipping normalization.")
        return raw_data

    logging.info(f"{prefix} Found {len(groups)} pallet group(s). Checking for misplaced values...")

    # --- Normalize: pull misplaced values up to anchor ---
    rows_to_remove = set()
    pull_up_count = 0

    for group in groups:
        anchor_idx = group['anchor_idx']
        anchor_row = raw_data[anchor_idx]

        for member_idx in group['member_indices']:
            if member_idx == anchor_idx:
                continue  # Skip the anchor itself

            member_row = raw_data[member_idx]

            # Is this member row misplaced? (has distributable values but NOT a pallet row)
            if not _is_pallet_row(member_row) and _has_distributable_value(member_row):
                # Track details for UI warning
                moved_details = []
                # Pull up each distributable column value to the anchor
                for col in resolved_cols:
                    val = member_row.get(col)
                    if val is not None:
                        try:
                            val_dec = val if isinstance(val, decimal.Decimal) else decimal.Decimal(str(val))
                            if val_dec != 0:
                                # Add to anchor (sum if anchor already has a value)
                                existing = anchor_row.get(col)
                                if existing is not None:
                                    try:
                                        existing_dec = existing if isinstance(existing, decimal.Decimal) else decimal.Decimal(str(existing))
                                        anchor_row[col] = existing_dec + val_dec
                                    except (decimal.InvalidOperation, ValueError, TypeError):
                                        anchor_row[col] = val_dec
                                else:
                                    anchor_row[col] = val_dec

                                # Handle pulling up col_cbm_raw specifically
                                raw_cbm_val = None
                                if col == 'col_cbm' and 'col_cbm_raw' in member_row:
                                    raw_cbm_val = member_row.get('col_cbm_raw')
                                    if raw_cbm_val:
                                        existing_raw = anchor_row.get('col_cbm_raw')
                                        if existing_raw:
                                            anchor_row['col_cbm_raw'] = f"{existing_raw} + {raw_cbm_val}"
                                        else:
                                            anchor_row['col_cbm_raw'] = str(raw_cbm_val)
                                    member_row['col_cbm_raw'] = None

                                # Clear the value from the filler row
                                member_row[col] = None
                                pull_up_count += 1
                                
                                # Track exact value moved for warning, use raw formula for CBM if available
                                if col == 'col_cbm' and raw_cbm_val:
                                    moved_details.append(f"CBM='{raw_cbm_val}'")
                                else:
                                    moved_details.append(f"{col}={val}")
                        except (decimal.InvalidOperation, ValueError, TypeError):
                            pass  # Can't convert — leave it

                logging.debug(f"{prefix} Pulled up values from row {member_idx} to anchor row {anchor_idx}")
                
                # Emit a user-visible warning via the monitor with exact values
                if monitor and moved_details:
                    po_val = anchor_row.get('col_po', 'Unknown PO')
                    item_val = anchor_row.get('col_item', 'Unknown Item')
                    moved_str = ", ".join(moved_details)
                    monitor.log_warning(
                        f"Auto-corrected misplaced data: Pulled up [{moved_str}] to Pallet Row (PO: {po_val}, Item: {item_val}). Please verify."
                    )

                # If filler row now has NO basis (pcs), mark for removal
                if not _has_basis_value(member_row):
                    rows_to_remove.add(member_idx)

    # --- Remove empty filler rows ---
    if rows_to_remove:
        result = [row for i, row in enumerate(raw_data) if i not in rows_to_remove]
        logging.info(f"{prefix} Normalization complete. Pulled up {pull_up_count} misplaced row(s), removed {len(rows_to_remove)} empty filler row(s). {len(result)} rows remain.")
    else:
        result = raw_data
        logging.info(f"{prefix} Normalization complete. Pulled up {pull_up_count} misplaced row(s), no rows removed. {len(result)} rows remain.")

    return result


# distribute_values function remains unchanged...
def distribute_values(
    raw_data: List[Dict[str, Any]],
    columns_to_distribute: List[str],
    basis_column: str
) -> List[Dict[str, Any]]:
    """
    Distributes values in specified columns based on proportions in the basis column.
    Operates on the input raw_data (which might have pre-calculated CBM).
    Handles pre-calculated CBM decimals correctly. Modifies data in place.
    """
    prefix = "[distribute_values]"
    logging.debug(f"{prefix} Starting value distribution process.")

    if not raw_data:
        logging.warning(f"{prefix} Received empty raw_data list. Skipping distribution.")
        return []

    processed_data = raw_data

    # --- Find Basis Column Canonical Name ---
    candidate_basis = basis_column
    if not candidate_basis.startswith('col_'):
        if basis_column == 'pcs':
             candidate_basis = 'col_qty_pcs'
        elif basis_column == 'sqft':
             candidate_basis = 'col_qty_sf'
        else:
             candidate_basis = f"col_{basis_column}"
    
    # Check if basis exists anywhere
    if not any(candidate_basis in row for row in processed_data):
        logging.error(f"{prefix} Basis column '{basis_column}' (or '{candidate_basis}') not found in any row. Cannot distribute.")
        raise ProcessingError(f"Basis column '{basis_column}' not found for distribution.")
    
    basis_column = candidate_basis

    valid_columns_to_distribute = []
    if columns_to_distribute:
        for col in columns_to_distribute:
            target_col = col
            if not target_col.startswith('col_'):
                if target_col == 'net': target_col = 'col_net'
                elif target_col == 'gross': target_col = 'col_gross'
                elif target_col == 'cbm': target_col = 'col_cbm'
                elif target_col == 'sqft': target_col = 'col_qty_sf'
                elif target_col == 'pcs': target_col = 'col_qty_pcs'
                elif target_col == 'amount': target_col = 'col_amount'
                elif target_col == 'pallet_count': target_col = 'col_pallet_count'
                else: target_col = f"col_{target_col}"

            # Only add to valid if it appears in at least one row
            if any(target_col in row for row in processed_data):
                valid_columns_to_distribute.append(target_col)
            else:
                logging.warning(f"{prefix} Column '{col}' (mapped to '{target_col}') not found in any row. Skipping.")
    else:
        logging.info(f"{prefix} No columns specified in 'columns_to_distribute' list. Skipping distribution.")
        return processed_data

    if not valid_columns_to_distribute:
         logging.warning(f"{prefix} No valid columns found to perform distribution on. Requested: {columns_to_distribute}")
         return processed_data

    num_rows = len(processed_data)
    logging.info(f"{prefix} Starting value distribution for columns: {valid_columns_to_distribute} based on '{basis_column}' ({num_rows} rows).")

    # Pre-convert basis values to Decimal
    basis_values_dec: List[Optional[decimal.Decimal]] = [
        _convert_to_decimal(row.get(basis_column), f"{prefix} basis column '{basis_column}' row index {i}")
        for i, row in enumerate(processed_data)
    ]
    logging.debug(f"{prefix} Pre-converted basis values (first 10): {basis_values_dec[:10]}")

    # --- Process each column ---
    for col_name in valid_columns_to_distribute:
        logging.info(f"{prefix} Processing column for distribution: '{col_name}'")

        # Pre-convert original values for the column being distributed
        current_col_values_dec: List[Optional[decimal.Decimal]] = [
             val if isinstance((val := row.get(col_name)), decimal.Decimal)
             else _convert_to_decimal(val, f"{prefix} column '{col_name}' row index {i}")
             for i, row in enumerate(processed_data)
        ]

        # Initialize processed list for this column
        processed_col_values: List[Optional[decimal.Decimal]] = [None] * num_rows

        i = 0
        while i < num_rows:
            current_val_dec = current_col_values_dec[i]
            log_row_context = f"{prefix} Col '{col_name}', Row index {i}"

            # --- Case 1: Found a non-None, non-zero value to potentially distribute ---
            if current_val_dec is not None and current_val_dec != decimal.Decimal(0):
                processed_col_values[i] = current_val_dec

                # --- Look ahead for the distribution block ---
                j = i + 1 
                distribution_rows_indices = []
                while j < num_rows:
                     next_original_val_dec = current_col_values_dec[j]
                     if next_original_val_dec is not None and next_original_val_dec != decimal.Decimal(0):
                          break

                     basis_for_j = basis_values_dec[j]
                     if basis_for_j is not None:
                          distribution_rows_indices.append(j)
                     else:
                          distribution_rows_indices.append(j)
                          logging.warning(f"{log_row_context}: Lookahead index {j} has MISSING basis. Will assign 0 later.")
                     j += 1

                if distribution_rows_indices:
                    block_indices = [i] + distribution_rows_indices
                    total_basis_in_block = decimal.Decimal(0)
                    indices_with_valid_basis = []

                    for k in block_indices:
                        basis_val = basis_values_dec[k]
                        if basis_val is not None and basis_val > 0:
                            total_basis_in_block += basis_val
                            indices_with_valid_basis.append(k)

                    if total_basis_in_block > 0 and indices_with_valid_basis:
                         distributed_sum_check = decimal.Decimal(0)
                         dist_precision = CBM_DECIMAL_PLACES if col_name == 'col_cbm' else DEFAULT_DIST_PRECISION

                         num_valid_indices = len(indices_with_valid_basis)
                         
                         if num_valid_indices == 1:
                             k = indices_with_valid_basis[0]
                             processed_col_values[k] = current_val_dec.quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                             distributed_sum_check = processed_col_values[k]
                         else:
                             for k in indices_with_valid_basis[:-1]:
                                 basis_val = basis_values_dec[k]
                                 proportion = basis_val / total_basis_in_block
                                 distributed_value = (current_val_dec * proportion).quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                                 processed_col_values[k] = distributed_value
                                 distributed_sum_check += distributed_value
                             
                             last_idx = indices_with_valid_basis[-1]
                             remainder = current_val_dec - distributed_sum_check
                             processed_col_values[last_idx] = remainder.quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                             distributed_sum_check += processed_col_values[last_idx]

                         for k in block_indices:
                             if k not in indices_with_valid_basis:
                                 if processed_col_values[k] is None:
                                     processed_col_values[k] = decimal.Decimal(0)

                         tolerance = dist_precision / decimal.Decimal(2)
                         diff = abs(distributed_sum_check - current_val_dec)
                         if not diff <= tolerance:
                              logging.warning(f"{log_row_context}: Distribution Check potentially FAILED for block. Diff: {diff:.10f}")

                    else:
                        for k in distribution_rows_indices:
                            if processed_col_values[k] is None:
                                processed_col_values[k] = decimal.Decimal(0)

                    i = j
                else:
                    i += 1

            # --- Case 2: Current original value is None or zero ---
            else:
                if processed_col_values[i] is None:
                    processed_col_values[i] = decimal.Decimal(0)
                i += 1

        # Push calculated values back into row dicts
        for idx, row in enumerate(processed_data):
             if processed_col_values[idx] is not None:
                  row[col_name] = processed_col_values[idx]

    logging.info(f"{prefix} Value distribution processing COMPLETED for all requested columns.")
    return processed_data

# *** Standard Aggregation Function (MODIFIED to handle SQFT, AMOUNT, and DESCRIPTION key) ***
def aggregate_standard_by_po_item_price(
    processed_data: List[Dict[str, Any]],
    global_aggregation_map: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]]
) -> Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]]:
    """
    STANDARD Aggregation: Aggregates 'sqft' AND 'amount' values based on unique
    combinations of 'po', 'item', 'unit' price, AND 'description'.
    Updates the global_aggregation_map in place.
    """
    aggregated_results = global_aggregation_map
    required_cols = ['col_po', 'col_item', 'col_unit_price', 'col_qty_sf', 'col_amount']
    prefix = "[aggregate_standard]"

    logging.debug(f"{prefix} Updating global STANDARD aggregation (SQFT & Amount by PO/Item/Price/Desc) with new table data.")
    logging.debug(f"{prefix} Size of global map BEFORE processing this table: {len(aggregated_results)}")

    if not processed_data:
        logging.info(f"{prefix} No data rows found in this table. Global map unchanged.")
        return aggregated_results

    # Check for required columns existing in at least one row
    missing_cols = [col for col in required_cols if not any(col in row for row in processed_data)]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform STANDARD aggregation: Missing required columns {missing_cols}. Skipping this table.")
        return aggregated_results

    has_description_col = any('col_desc' in row for row in processed_data)
    if not has_description_col:
        logging.info(f"{prefix} 'col_desc' column not found or is invalid. Will use None for description keys.")

    num_rows = len(processed_data)
    logging.info(f"{prefix} Processing {num_rows} rows for STANDARD aggregation (SQFT & Amount by PO/Item/Price/Desc).")

    rows_processed_this_table = 0
    successful_conversions_sqft = 0
    successful_conversions_amount = 0

    for i, row in enumerate(processed_data):
        rows_processed_this_table += 1
        log_row_context = f"{prefix} Table Row index {i}"
        
        po_val, item_val = row.get('col_po'), row.get('col_item')
        unit_price_raw, sqft_raw, amount_raw = row.get('col_unit_price'), row.get('col_qty_sf'), row.get('col_amount')
        desc_raw = row.get('col_desc') if has_description_col else None

        logging.debug(f"{log_row_context}: Raw values - PO='{po_val}', Item='{item_val}', Price='{unit_price_raw}', Desc='{desc_raw}', SQFT='{sqft_raw}', Amount='{amount_raw}'")

        # Prepare key components
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw # Keep None as None
        description_key = description_key if description_key else None # Ensure empty strings become None

        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"
        # Description key can be None

        # Convert price to Decimal for the key
        price_dec = DataConverter.convert_to_decimal(unit_price_raw, f"{log_row_context} price")

        # UPDATED Key: (PO, Item, Price, Description)
        key = (po_key, item_key, price_dec, description_key)
        logging.debug(f"{log_row_context}: Generated Key Tuple = {key}")


        # Convert SQFT and Amount to Decimal for summation
        sqft_dec = _convert_to_decimal(sqft_raw, f"{log_row_context} SQFT")
        if sqft_dec is None:
             # logging.debug(f"{log_row_context}: SQFT value '{sqft_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
             sqft_dec = decimal.Decimal(0)
        else:
             successful_conversions_sqft +=1

        amount_dec = _convert_to_decimal(amount_raw, f"{log_row_context} Amount")
        if amount_dec is None:
            # logging.debug(f"{log_row_context}: Amount value '{amount_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
            amount_dec = decimal.Decimal(0)
        else:
            successful_conversions_amount +=1

        # logging.debug(f"{log_row_context}: Converted values - SQFT='{sqft_dec}', Amount='{amount_dec}'") # Reduced verbosity

        # --- Add to the global aggregate sums (SQFT, Amount, Net) ---
        current_sums = aggregated_results.get(key, {'sqft_sum': decimal.Decimal(0), 'amount_sum': decimal.Decimal(0), 'net_sum': decimal.Decimal(0)})

        # Update the sums
        current_sums['sqft_sum'] += sqft_dec
        current_sums['amount_sum'] += amount_dec
        
        # Aggregate Net Weight if present
        net_val = row.get('col_net')
        if net_val is not None:
            net_dec = _convert_to_decimal(net_val, f"{log_row_context} Net")
            if net_dec is not None:
                current_sums['net_sum'] += net_dec
        
        # Track col_cbm_raw specifically
        raw_cbm_val = row.get('col_cbm_raw')
        if raw_cbm_val:
            existing_raw = current_sums.get('col_cbm_raw')
            if existing_raw:
                current_sums['col_cbm_raw'] = f"{existing_raw} + {raw_cbm_val}"
            else:
                current_sums['col_cbm_raw'] = str(raw_cbm_val)

        # Store the updated dictionary back into the global map
        aggregated_results[key] = current_sums
        # logging.debug(f"{log_row_context}: Global sums for key {key} AFTER add = {aggregated_results[key]}") # Reduced verbosity


    logging.info(f"{prefix} Finished processing {rows_processed_this_table} rows.")
    logging.info(f"{prefix} SQFT values successfully converted/defaulted for {successful_conversions_sqft} rows.")
    logging.info(f"{prefix} Amount values successfully converted/defaulted for {successful_conversions_amount} rows.")
    logging.info(f"{prefix} Global standard aggregation map size: {len(aggregated_results)}")
    # logging.debug(f"{prefix} Global Standard Aggregated Results (End of Table):\n{pprint.pformat(aggregated_results)}") # Keep DEBUG for detailed tracing if needed
    return aggregated_results


# *** Custom Aggregation Function (MODIFIED to include DESCRIPTION key) ***
def aggregate_custom_by_po_item(
    processed_data: List[Dict[str, Any]],
    # UPDATED TYPE HINT: Key is now 4 elements (PO, Item, None, Description)
    global_custom_aggregation_map: Dict[Tuple[Any, Any, None, Optional[str]], Dict[str, decimal.Decimal]]
) -> Dict[Tuple[Any, Any, None, Optional[str]], Dict[str, decimal.Decimal]]:
    """
    CUSTOM Aggregation: Aggregates 'sqft' and 'amount' values based on unique
    combinations of 'po', 'item', AND 'description'. Uses a 4-element key
    (PO, Item, None, Description) for structural consistency with standard aggregation.
    Updates the global_custom_aggregation_map in place.

    Args:
        processed_data: Dictionary representing the data of the current table.
        global_custom_aggregation_map: The dictionary holding the cumulative custom
                                       aggregation results.
                                       Key: (po, item, None, description)
                                       Value: Dict{'sqft_sum': Decimal, 'amount_sum': Decimal}.

    Returns:
        The updated global_custom_aggregation_map.
    """
    aggregated_results = global_custom_aggregation_map
    # Required columns for this aggregation (Description is optional)
    required_cols = ['col_po', 'col_item', 'col_qty_sf', 'col_amount']
    prefix = "[aggregate_custom]"

    logging.debug(f"{prefix} Updating global CUSTOM aggregation (SQFT & Amount by PO/Item/Desc) with new table data.")
    logging.debug(f"{prefix} Size of global custom map BEFORE processing this table: {len(aggregated_results)}")
    # Check for required columns existing in at least one row
    missing_cols = [col for col in required_cols if not any(col in row for row in processed_data)]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform full CUSTOM aggregation: Missing required columns {missing_cols}. Proceeding cautiously.")

    has_description_col = any('col_desc' in row for row in processed_data)
    if not has_description_col:
        logging.info(f"{prefix} 'col_desc' column not found or is invalid. Will use None for description keys.")

    num_rows = len(processed_data)
    if num_rows == 0:
        logging.info(f"{prefix} No data rows found in this table. Global custom aggregation map remains unchanged.")
        return aggregated_results

    logging.info(f"{prefix} Processing {num_rows} rows from this table to update global CUSTOM aggregation (by PO/Item/Desc).")

    # --- Iterate and Aggregate ---
    rows_processed_this_table = 0
    successful_conversions_sqft = 0
    successful_conversions_amount = 0

    for i, row in enumerate(processed_data):
        rows_processed_this_table += 1
        log_row_context = f"{prefix} Table Row index {i}"
        
        # Get raw values 
        po_val, item_val = row.get('col_po'), row.get('col_item')
        sqft_raw, amount_raw = row.get('col_qty_sf'), row.get('col_amount')
        desc_raw = row.get('col_desc') if has_description_col else None

        # Prepare the key components (Handle None, strip strings)
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw
        description_key = description_key if description_key else None # Ensure empty strings become None

        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"
        # Description key can be None

        # UPDATED Key: (PO, Item, None, Description) - Move description to index 3 to match standard
        key = (po_key, item_key, None, description_key)

        # Convert SQFT to Decimal for summation (default to 0 if fails/None)
        sqft_dec = _convert_to_decimal(sqft_raw, f"{log_row_context} SQFT")
        if sqft_dec is None:
             sqft_dec = decimal.Decimal(0)
        else:
             successful_conversions_sqft +=1

        # Convert Amount to Decimal for summation (default to 0 if fails/None)
        amount_dec = _convert_to_decimal(amount_raw, f"{log_row_context} Amount")
        if amount_dec is None:
            amount_dec = decimal.Decimal(0)
        else:
            successful_conversions_amount +=1

        # --- Add to the global aggregate sums ---
        current_sums = aggregated_results.get(key, {'sqft_sum': decimal.Decimal(0), 'amount_sum': decimal.Decimal(0), 'net_sum': decimal.Decimal(0)})

        # Update the sums
        current_sums['sqft_sum'] += sqft_dec
        current_sums['amount_sum'] += amount_dec
        
        # Aggregate Net Weight if present
        net_val = row.get('col_net')
        if net_val is not None:
            net_dec = _convert_to_decimal(net_val, f"{log_row_context} Net")
            if net_dec is not None:
                current_sums['net_sum'] += net_dec
        
        # Track col_cbm_raw specifically
        raw_cbm_val = row.get('col_cbm_raw')
        if raw_cbm_val:
            existing_raw = current_sums.get('col_cbm_raw')
            if existing_raw:
                current_sums['col_cbm_raw'] = f"{existing_raw} + {raw_cbm_val}"
            else:
                current_sums['col_cbm_raw'] = str(raw_cbm_val)

        # Store the updated dictionary back into the global map
        aggregated_results[key] = current_sums

    # --- Log summary for this table's contribution ---
    logging.info(f"{prefix} Finished processing {rows_processed_this_table} rows for this table.")
    logging.info(f"{prefix} SQFT values successfully converted/defaulted for {successful_conversions_sqft} rows.")
    logging.info(f"{prefix} Amount values successfully converted/defaulted for {successful_conversions_amount} rows.")
    logging.info(f"{prefix} Global custom aggregation map now contains {len(aggregated_results)} unique (PO, Item, None, Description) keys.")

    return aggregated_results


def calculate_leather_summary(processed_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Calculates the leather summary (PCS, SQFT, Net, Gross, Pallet Count) per leather type.
    Iterates through rows to sum values for each leather type found in 'description' or 'desc'.
    BUFFALO = rows containing "BUFFALO" in description
    COW = rows NOT containing "BUFFALO" (all other leather)
    """
    # Initialize summary structure with default 0s
    summary = {
        'BUFFALO': {'col_qty_pcs': 0, 'col_qty_sf': decimal.Decimal(0), 'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0), 'col_cbm': decimal.Decimal(0), 'col_pallet_count': 0},
        'COW': {'col_qty_pcs': 0, 'col_qty_sf': decimal.Decimal(0), 'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0), 'col_cbm': decimal.Decimal(0), 'col_pallet_count': 0}
    }

    if not processed_data:
        return summary

    for row in processed_data:
        desc = str(row.get('col_desc', "")).upper() if row.get('col_desc') else ""
        
        # BUFFALO = contains "BUFFALO", COW = everything else (non-buffalo leather)
        leather_type = 'BUFFALO' if "BUFFALO" in desc else 'COW'
        
        # Sum PCS
        try:
            val = row.get('col_qty_pcs')
            summary[leather_type]['col_qty_pcs'] += int(float(val)) if val else 0
        except (ValueError, TypeError): pass

        # Sum SQFT
        try:
            val = row.get('col_qty_sf')
            summary[leather_type]['col_qty_sf'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError): pass

        # Sum Net
        try:
            val = row.get('col_net')
            summary[leather_type]['col_net'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError): pass

        # Sum Gross
        try:
            val = row.get('col_gross')
            summary[leather_type]['col_gross'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError): pass

        # Sum CBM
        try:
            val = row.get('col_cbm')
            summary[leather_type]['col_cbm'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError): pass

        # Sum Pallet Count (if available per row)
        try:
            val = row.get('col_pallet_count')
            summary[leather_type]['col_pallet_count'] += int(float(val)) if val else 0
        except (ValueError, TypeError): pass

    return summary


def aggregate_per_po_with_pallets(processed_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Aggregates data by PO and Item, summing pallet, pcs, sqft, amount, net, gross, cbm.
    Groups rows that share the same (PO, Item) combination.

    Args:
        processed_data: List of row dicts with col_* keys.

    Returns:
        A list of aggregated records keyed by (PO, Item) with col_* totals.
    """
    if not isinstance(processed_data, list) or not processed_data:
        return []

    aggregation_map = {}

    for row in processed_data:
        # --- Validate PO ---
        po_val = row.get('col_po')
        if po_val is None:
            continue
        po = str(po_val).strip()
        if not po:
            continue

        # --- Validate Item ---
        item_val = row.get('col_item')
        item = str(item_val).strip() if item_val is not None else ""

        key = (po, item)

        if key not in aggregation_map:
            aggregation_map[key] = {
                'col_desc': '',
                'col_qty_pcs': 0,
                'col_qty_sf': decimal.Decimal(0),
                'col_amount': decimal.Decimal(0),
                'col_pallet_count': 0,
                'col_net': decimal.Decimal(0),
                'col_gross': decimal.Decimal(0),
                'col_cbm': decimal.Decimal(0),
                'col_cbm_raw': ''
            }

        # Capture description from first row (all rows in same PO+Item group share desc)
        if not aggregation_map[key]['col_desc']:
            desc_val = row.get('col_desc')
            if desc_val:
                aggregation_map[key]['col_desc'] = str(desc_val).strip()

        # Concatenate col_cbm_raw strings
        cbm_raw_val = row.get('col_cbm_raw')
        if cbm_raw_val:
            cbm_raw_str = str(cbm_raw_val).strip()
            if cbm_raw_str:
                existing_raw = aggregation_map[key]['col_cbm_raw']
                if existing_raw:
                    # Avoid duplicating the same formula "1.2*0.8*1 + 1.2*0.8*1" if they are identical
                    # But for now, we just join them to show what was aggregated
                    if cbm_raw_str not in str(existing_raw).split(" + "):
                        aggregation_map[key]['col_cbm_raw'] = str(aggregation_map[key]['col_cbm_raw']) + f" + {cbm_raw_str}"
                else:
                    aggregation_map[key]['col_cbm_raw'] = cbm_raw_str

        # Sum sqft
        sqft_val = row.get('col_qty_sf')
        if sqft_val is not None:
             converted = _convert_to_decimal(sqft_val)
             if converted:
                 aggregation_map[key]['col_qty_sf'] += converted

        # Sum amount
        amount_val = row.get('col_amount')
        if amount_val is not None:
             converted = _convert_to_decimal(amount_val)
             if converted:
                 aggregation_map[key]['col_amount'] += converted

        # Sum pallet_count — prefer col_pallet_count_raw (integer backup)
        # because col_pallet_count may already be an "x-y" format string
        # after main.py's pallet order conversion.
        pallet_val = row.get('col_pallet_count_raw', row.get('col_pallet_count'))
        if pallet_val is not None:
             try:
                 aggregation_map[key]['col_pallet_count'] += int(float(pallet_val))
             except (ValueError, TypeError): pass

        # Sum net
        net_val = row.get('col_net')
        if net_val is not None:
             converted = _convert_to_decimal(net_val)
             if converted:
                 aggregation_map[key]['col_net'] += converted

        # Sum gross
        gross_val = row.get('col_gross')
        if gross_val is not None:
             converted = _convert_to_decimal(gross_val)
             if converted:
                 aggregation_map[key]['col_gross'] += converted

        # Sum cbm
        cbm_val = row.get('col_cbm')
        if cbm_val is not None:
             converted = _convert_to_decimal(cbm_val)
             if converted:
                 aggregation_map[key]['col_cbm'] += converted

        # Sum pcs
        pcs_val = row.get('col_qty_pcs')
        if pcs_val is not None:
             try:
                 aggregation_map[key]['col_qty_pcs'] += int(float(pcs_val))
             except (ValueError, TypeError): pass

    # Convert to list of dicts
    result = []
    for (po, item), data in aggregation_map.items():
        result.append({
            'col_po': po,
            'col_item': item,
            'col_desc': data['col_desc'],
            'col_qty_pcs': data['col_qty_pcs'],
            'col_qty_sf': data['col_qty_sf'],
            'col_amount': data['col_amount'],
            'col_pallet_count': data['col_pallet_count'],
            'col_net': data['col_net'],
            'col_gross': data['col_gross'],
            'col_cbm': data['col_cbm'],
            'col_cbm_raw': data['col_cbm_raw']
        })

    # Sort by PO, then by Item for consistent output
    result.sort(key=lambda x: (x['col_po'], x['col_item']))

    logging.info(f"[aggregate_per_po_with_pallets] Aggregated {len(processed_data)} rows into {len(result)} unique PO+Item combinations.")

    return result


def calculate_weight_summary(processed_data: List[Dict[str, Any]]) -> Dict[str, decimal.Decimal]:
    """
    Calculates the weight summary (Net Weight and Gross Weight).
    
    Args:
        processed_data: List of Dictionary representing the data of the current table.
        
    Returns:
        Dictionary containing 'net' and 'gross' weights.
    """
    summary = {'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0)}
    
    if not processed_data:
        return summary
        
    for row in processed_data:
        net_val = _convert_to_decimal(row.get('col_net'))
        if net_val is not None:
             summary['col_net'] += net_val
             
        gross_val = _convert_to_decimal(row.get('col_gross'))
        if gross_val is not None:
             summary['col_gross'] += gross_val
             
    return summary

def calculate_pallet_summary(processed_data: List[Dict[str, Any]]) -> int:
    """
    Calculates the total pallet count for the table.
    
    Args:
        processed_data: List of Dictionaries representing the data of the current table.
        
    Returns:
        Total pallet count as integer.
    """
    total_pallets = 0
    
    if not processed_data:
        return 0
        
    for row in processed_data:
        val = row.get('col_pallet_count')
        if val is not None:
            try:
                total_pallets += int(float(val))
            except (ValueError, TypeError):
                pass
                
    return total_pallets

def calculate_footer_totals(processed_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Calculates totals for footer fields based on processed data.
    Returns a dictionary with keys matching the expected footer data structure.
    """
    totals = {
        "col_qty_pcs": 0,
        "col_qty_sf": decimal.Decimal(0),
        "col_net": decimal.Decimal(0),
        "col_gross": decimal.Decimal(0),
        "col_cbm": decimal.Decimal(0),
        "col_amount": decimal.Decimal(0),
        "col_pallet_count": 0
    }
    
    if not processed_data:
        return totals

    def safe_add_decimal(key, value):
        if value is not None:
            try:
                val_str = str(value).replace(',', '')
                totals[key] += decimal.Decimal(val_str)
            except (decimal.InvalidOperation, ValueError, TypeError):
                pass

    def safe_add_int(key, value):
        if value is not None:
            try:
                val_str = str(value).replace(',', '')
                totals[key] += int(float(val_str))
            except (ValueError, TypeError):
                pass

    # Sum row by row
    for row in processed_data:
        safe_add_int('col_qty_pcs', row.get('col_qty_pcs'))
        safe_add_decimal('col_qty_sf', row.get('col_qty_sf'))
        safe_add_decimal('col_net', row.get('col_net'))
        safe_add_decimal('col_gross', row.get('col_gross'))
        safe_add_decimal('col_cbm', row.get('col_cbm'))
        safe_add_decimal('col_amount', row.get('col_amount'))
        safe_add_int('col_pallet_count', row.get('col_pallet_count'))

    return totals


def format_aggregation_as_list(
    aggregation_map: Dict[Tuple, Dict[str, decimal.Decimal]],
    mode: str = 'standard'
) -> List[Dict[str, Any]]:
    """
    Converts the internal tuple-keyed aggregation map into a clean list of dictionaries
    suitable for JSON output. Removes tuple keys and uses 'col_' prefixed keys for all fields.

    Args:
        aggregation_map: The dictionary holding the aggregation results.
        mode: 'standard' or 'custom' indicating the key structure.

    Returns:
        A list of dictionaries, each representing an aggregated row.
    """
    flattened_list = []
    
    for key_tuple, values in aggregation_map.items():
        row_dict = {}
        
        # Extract values from the tuple key based on mode
        if mode == 'standard':
            # Key: (PO, Item, Price, Desc)
            # Mapping: Index 0->col_po, 1->col_item, 2->col_unit_price, 3->col_desc
            if len(key_tuple) >= 4:
                row_dict['col_po'] = str(key_tuple[0]) if key_tuple[0] is not None else ""
                row_dict['col_item'] = str(key_tuple[1]) if key_tuple[1] is not None else ""
                row_dict['col_unit_price'] = str(key_tuple[2]) if key_tuple[2] is not None else ""
                row_dict['col_desc'] = str(key_tuple[3]) if key_tuple[3] is not None else ""
            else:
                # Fallback for unexpected key length
                row_dict['col_po'] = str(key_tuple[0]) if len(key_tuple) > 0 else ""
                row_dict['col_item'] = str(key_tuple[1]) if len(key_tuple) > 1 else ""
                
        elif mode == 'custom':
            # Key: (PO, Item, None, Desc)
            # Mapping: Index 0->col_po, 1->col_item, 3->col_desc (Index 2 is None/Ignored)
            if len(key_tuple) >= 4:
                row_dict['col_po'] = str(key_tuple[0]) if key_tuple[0] is not None else ""
                row_dict['col_item'] = str(key_tuple[1]) if key_tuple[1] is not None else ""
                row_dict['col_desc'] = str(key_tuple[3]) if key_tuple[3] is not None else ""
            else:
                 # Fallback
                row_dict['col_po'] = str(key_tuple[0]) if len(key_tuple) > 0 else ""
                row_dict['col_item'] = str(key_tuple[1]) if len(key_tuple) > 1 else ""

        # Extract aggregated values
        sqft_sum = values.get('sqft_sum', decimal.Decimal(0))
        amount_sum = values.get('amount_sum', decimal.Decimal(0))
        net_sum = values.get('net_sum', decimal.Decimal(0))
        
        row_dict['col_qty_sf'] = float(sqft_sum) if isinstance(sqft_sum, decimal.Decimal) else sqft_sum
        row_dict['col_amount'] = float(amount_sum) if isinstance(amount_sum, decimal.Decimal) else amount_sum
        row_dict['col_net'] = float(net_sum) if isinstance(net_sum, decimal.Decimal) else net_sum
        
        if 'col_cbm_raw' in values:
            row_dict['col_cbm_raw'] = values['col_cbm_raw']
            
        flattened_list.append(row_dict)
        
    return flattened_list