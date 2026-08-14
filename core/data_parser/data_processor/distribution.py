import logging
import decimal
from typing import List, Dict, Any, Optional
from ..util.converters import DataConverter
from ..validation import DataValidationError
from core.utils.pipeline_monitor import PipelineMonitor

_convert_to_decimal = DataConverter.convert_to_decimal
CBM_DECIMAL_PLACES = decimal.Decimal('0.01')
DEFAULT_DIST_PRECISION = decimal.Decimal('0.0001')

# Custom exception for data processing errors.
# Subclasses DataValidationError so callers catching that base class see these too.
class ProcessingError(DataValidationError):
    """Raised when a pipeline step cannot proceed due to bad/missing data."""
    pass

def inject_net_weight_pricing(tables: List[List[Dict[str, Any]]], global_unit_price: float) -> List[List[Dict[str, Any]]]:
    """
    For 'net' pricing mode: injects col_unit_price, col_amount, and col_qty_sf
    from col_net and the user-provided global unit price.
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
                    if row.get('col_qty_sf') is None:
                        row['col_qty_sf'] = 0.0
                    injected_count += 1

    logging.info(f"{prefix} Injected unit_price={price} into {injected_count} rows across {len(tables)} tables.")
    return tables


def normalize_by_pallet_anchor(
    raw_data: List[Dict[str, Any]],
    columns_to_normalize: List[str],
    basis_column: str = 'col_qty_pcs',
    pallet_column: str = 'col_pallet_count',
    monitor: Optional[PipelineMonitor] = None
) -> List[Dict[str, Any]]:
    """
    Uses col_pallet_count as anchor to fix misplaced distributable values.
    """
    prefix = "[normalize_by_pallet_anchor]"

    if not raw_data:
        return []

    for col in columns_to_normalize:
        if not col.startswith('col_'):
            raise ValueError(f"Normalization Error: Column name '{col}' must start with 'col_'.")
    
    if not basis_column.startswith('col_'):
        raise ValueError(f"Normalization Error: Basis column '{basis_column}' must start with 'col_'.")

    resolved_cols = list(columns_to_normalize)
    resolved_basis = basis_column

    has_pallet_col = any(pallet_column in row for row in raw_data)
    if not has_pallet_col:
        logging.info(f"{prefix} No '{pallet_column}' found in data. Skipping normalization.")
        return raw_data

    def _is_pallet_row(row: Dict[str, Any]) -> bool:
        val = row.get(pallet_column)
        if val is None:
            return False
        try:
            return decimal.Decimal(str(val)) >= 1
        except (decimal.InvalidOperation, ValueError, TypeError):
            return False

    def _has_distributable_value(row: Dict[str, Any]) -> bool:
        for col in resolved_cols:
            val = row.get(col)
            if val is not None:
                try:
                    if decimal.Decimal(str(val)) != 0:
                        return True
                except (decimal.InvalidOperation, ValueError, TypeError):
                    if str(val).strip():
                        return True
        return False

    def _has_basis_value(row: Dict[str, Any]) -> bool:
        val = row.get(resolved_basis)
        if val is None:
            return False
        try:
            return decimal.Decimal(str(val)) > 0
        except (decimal.InvalidOperation, ValueError, TypeError):
            return False

    groups = []
    current_group = None

    for i, row in enumerate(raw_data):
        if _is_pallet_row(row):
            current_group = {'anchor_idx': i, 'member_indices': [i]}
            groups.append(current_group)
        elif current_group is not None:
            current_group['member_indices'].append(i)
        else:
            if _has_distributable_value(row):
                row_num = row.get('_row_num')
                row_num_str = f" (Row {row_num})" if row_num else ""
                po_val = row.get('col_po', 'Unknown PO')
                item_val = row.get('col_item', 'Unknown Item')
                error_msg = (
                    f"Data Validation Error{row_num_str}: Orphaned data found at PO [{po_val}] / Item [{item_val}]. "
                    f"Net Weight, Gross Weight, or CBM values exist before the first pallet anchor. "
                    f"Please ensure every pallet group starts with a valid Pallet Count."
                )
                if monitor:
                    monitor.log_process_item("Pallet Normalization", status="error", error=error_msg)
                raise DataValidationError(error_msg)

    if not groups:
        logging.info(f"{prefix} No pallet anchor rows found. Skipping normalization.")
        return raw_data

    logging.info(f"{prefix} Found {len(groups)} pallet group(s). Checking for misplaced values...")

    rows_to_remove = set()
    pull_up_count = 0

    for group in groups:
        anchor_idx = group['anchor_idx']
        anchor_row = raw_data[anchor_idx]

        for member_idx in group['member_indices']:
            if member_idx == anchor_idx:
                continue

            member_row = raw_data[member_idx]

            if not _is_pallet_row(member_row) and _has_distributable_value(member_row):
                moved_details = []
                for col in resolved_cols:
                    val = member_row.get(col)
                    if val is not None and val != 0:  # val is already Decimal from normalization
                        existing = anchor_row.get(col)
                        anchor_row[col] = (existing + val) if existing is not None else val

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

                        member_row[col] = None
                        pull_up_count += 1

                        if col == 'col_cbm' and raw_cbm_val:
                            moved_details.append(f"CBM='{raw_cbm_val}'")
                        else:
                            moved_details.append(f"{col}={val}")

                logging.debug(f"{prefix} Pulled up values from row {member_idx} to anchor row {anchor_idx}")
                
                if monitor and moved_details:
                    po_val = anchor_row.get('col_po', 'Unknown PO')
                    item_val = anchor_row.get('col_item', 'Unknown Item')
                    moved_str = ", ".join(moved_details)
                    monitor.log_warning(
                        f"Auto-corrected misplaced data: Pulled up [{moved_str}] to Pallet Row (PO: {po_val}, Item: {item_val}). Please verify."
                    )

                if not _has_basis_value(member_row):
                    rows_to_remove.add(member_idx)

    if rows_to_remove:
        result = [row for i, row in enumerate(raw_data) if i not in rows_to_remove]
        logging.info(f"{prefix} Normalization complete. Pulled up {pull_up_count} misplaced row(s), removed {len(rows_to_remove)} empty filler row(s). {len(result)} rows remain.")
    else:
        result = raw_data
        logging.info(f"{prefix} Normalization complete. Pulled up {pull_up_count} misplaced row(s), no rows removed. {len(result)} rows remain.")

    return result


def distribute_values(
    raw_data: List[Dict[str, Any]],
    columns_to_distribute: List[str],
    basis_column: str
) -> List[Dict[str, Any]]:
    """
    Distributes values in specified columns based on proportions in the basis column.
    """
    prefix = "[distribute_values]"
    logging.debug(f"{prefix} Starting value distribution process.")

    if not raw_data:
        logging.warning(f"{prefix} Received empty raw_data list. Skipping distribution.")
        return []

    processed_data = raw_data

    if not basis_column.startswith('col_'):
        raise ValueError(f"basis_column must be col_-prefixed, got '{basis_column}'. Fix the call site.")

    if not any(basis_column in row for row in processed_data):
        logging.error(f"{prefix} Basis column '{basis_column}' not found in any row. Cannot distribute.")
        raise ProcessingError(f"Basis column '{basis_column}' not found for distribution.")

    valid_columns_to_distribute = []
    if columns_to_distribute:
        for col in columns_to_distribute:
            if not col.startswith('col_'):
                raise ValueError(f"columns_to_distribute entry must be col_-prefixed, got '{col}'. Fix the call site.")
            if any(col in row for row in processed_data):
                valid_columns_to_distribute.append(col)
            else:
                logging.warning(f"{prefix} Column '{col}' not found in any row. Skipping.")
    else:
        logging.info(f"{prefix} No columns specified in 'columns_to_distribute' list. Skipping distribution.")
        return processed_data

    if not valid_columns_to_distribute:
        logging.warning(f"{prefix} No valid columns found to perform distribution on. Requested: {columns_to_distribute}")
        return processed_data

    num_rows = len(processed_data)
    logging.info(f"{prefix} Starting value distribution for columns: {valid_columns_to_distribute} based on '{basis_column}' ({num_rows} rows).")

    # Basis column is already Decimal from normalization — read directly
    basis_values_dec: List[Optional[decimal.Decimal]] = [
        row.get(basis_column) for row in processed_data
    ]
    logging.debug(f"{prefix} Pre-converted basis values (first 10): {basis_values_dec[:10]}")

    for col_name in valid_columns_to_distribute:
        logging.info(f"{prefix} Processing column for distribution: '{col_name}'")

        # Values already Decimal from normalization — read directly
        current_col_values_dec: List[Optional[decimal.Decimal]] = [
            row.get(col_name) for row in processed_data
        ]

        processed_col_values: List[Optional[decimal.Decimal]] = [None] * num_rows

        i = 0
        while i < num_rows:
            current_val_dec = current_col_values_dec[i]
            log_row_context = f"{prefix} Col '{col_name}', Row index {i}"

            if current_val_dec is not None and current_val_dec != decimal.Decimal(0):
                processed_col_values[i] = current_val_dec

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

            else:
                if processed_col_values[i] is None:
                    processed_col_values[i] = decimal.Decimal(0)
                i += 1

        for idx, row in enumerate(processed_data):
             if processed_col_values[idx] is not None:
                  if col_name in row:
                       row[f"{col_name}_raw"] = row[col_name]
                  row[col_name] = processed_col_values[idx]

    logging.info(f"{prefix} Value distribution processing COMPLETED for all requested columns.")
    return processed_data
