import logging
import decimal
from typing import List, Dict, Any, Optional
from .util.converters import DataConverter
from .config import HEADER_SEARCH_COL_RANGE, HEADERLESS_COLUMN_PATTERNS
from openpyxl.utils import get_column_letter
import re

# Set precision for Decimal calculations (consistent with data_processor)
decimal.getcontext().prec = 28

class DataValidationError(Exception):
    """User-facing validation error for missing required data.
    
    This exception is caught separately in the API layer to return
    a clean, human-readable error message without traceback noise.
    """
    pass


def _find_first_data_row(sheet, header_row: int, max_scan: int = 10) -> int:
    """Finds the first row index below header_row that contains actual table data."""
    for r in range(header_row + 1, min(header_row + 1 + max_scan, sheet.max_row + 1)):
        has_data = False
        for c in range(HEADER_SEARCH_COL_RANGE[0], HEADER_SEARCH_COL_RANGE[1] + 1):
            val = sheet.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                has_data = True
                break
        if has_data:
            return r
    return header_row + 1


def validate_no_duplicate_amount_columns(sheet, header_row: int, column_mapping: Optional[Dict[str, str]] = None):
    """
    Scans the header row and raises DataValidationError if multiple columns
    are mapped to 'col_amount' (via header aliases or data pattern matching).
    """
    from .sheet_parser import _ALIAS_REVERSE_LOOKUP

    # Find the first row below header that actually has data
    data_row = _find_first_data_row(sheet, header_row)

    # Collect columns that have already been mapped to other headers (not col_amount)
    mapped_cols = set()
    if column_mapping:
        for canonical, col_letter in column_mapping.items():
            if canonical != 'col_amount' and col_letter:
                mapped_cols.add(col_letter.upper())

    amount_cols = []
    alias_matched_cols = []
    
    # 1. First pass: check explicit header text alias mapping
    for col_num in range(HEADER_SEARCH_COL_RANGE[0], HEADER_SEARCH_COL_RANGE[1] + 1):
        cell = sheet.cell(row=header_row, column=col_num)
        val = str(cell.value or '').strip().upper()
        if val:
            col_letter = get_column_letter(col_num)
            if col_letter.upper() not in mapped_cols:
                candidates = _ALIAS_REVERSE_LOOKUP.get(val, [])
                if "col_amount" in candidates:
                    alias_matched_cols.append(col_num)
                    amount_cols.append(f"{get_column_letter(col_num)} ('{cell.value or '<empty>'}')")

    # 2. Second pass: check unrecognized/unmapped columns using pattern + left adjacent cell < 2 heuristic
    has_explicit_amount = len(alias_matched_cols) > 0
    for col_num in range(HEADER_SEARCH_COL_RANGE[0], HEADER_SEARCH_COL_RANGE[1] + 1):
        if col_num in alias_matched_cols:
            continue

        cell = sheet.cell(row=header_row, column=col_num)
        val = str(cell.value or '').strip().upper()

        # If we have an explicit alias match elsewhere, only perform the heuristic check
        # on truly headerless/empty columns. If no explicit alias match exists, we can
        # check any unrecognized column.
        if has_explicit_amount and val != "":
            continue

        col_letter = get_column_letter(col_num)
        if col_letter.upper() not in mapped_cols:
            amount_patterns = HEADERLESS_COLUMN_PATTERNS.get("col_amount", [])
            if amount_patterns:
                data_cell = sheet.cell(row=data_row, column=col_num)
                data_val = data_cell.value
                
                matches_pattern = False
                if data_val is not None:
                    # Robust handling of float stringification for pattern match
                    val_strs = [str(data_val).strip()]
                    try:
                        # Only format as 2-decimal if int/float/Decimal or contains '.'
                        if isinstance(data_val, (int, float, decimal.Decimal)) or '.' in str(data_val):
                            clean_str = str(data_val).replace(',', '').strip()
                            if re.match(r'^-?\d+(\.\d+)?$', clean_str):
                                num_val = float(clean_str)
                                val_strs.append(f"{num_val:.2f}")
                    except (ValueError, TypeError):
                        pass

                    for pattern in amount_patterns:
                        for val_str in val_strs:
                            try:
                                if re.match(pattern, val_str):
                                    matches_pattern = True
                                    break
                            except re.error:
                                continue
                        if matches_pattern:
                            break
                            
                # Left adjacent cell heuristic check (value < 2)
                has_left_unit_price = False
                if col_num > 1:
                    left_cell = sheet.cell(row=data_row, column=col_num - 1)
                    left_val = left_cell.value
                    left_num = None
                    if left_val is not None:
                        try:
                            left_num = float(str(left_val).replace(',', '').strip())
                        except ValueError:
                            pass
                    
                    if left_num is not None and 0 < left_num < 2:
                        # Ensure current cell also has a valid numeric value
                        if data_val is not None:
                            try:
                                float(str(data_val).replace(',', '').strip())
                                has_left_unit_price = True
                            except ValueError:
                                pass
                                
                if matches_pattern and has_left_unit_price:
                    amount_cols.append(f"{get_column_letter(col_num)} ('{cell.value or '<empty>'}')")
                
    if len(amount_cols) > 1:
        raise DataValidationError(
            f"Data Validation Error: Duplicate 'col_amount' columns detected on header row {header_row}: {', '.join(amount_cols)}. "
            f"Please ensure only one column is mapped to 'Amount'."
        )


def validate_col_level_discount(
    table_data: List[Dict[str, Any]], 
    table_id_str: Optional[str] = None, 
    monitor: Optional[Any] = None
) -> None:
    """
    Scans table_data for any rows where 'col_level' contains '折扣'.
    Logs a warning if discount rows are found.
    """
    if not table_data:
        return

    discount_rows = []
    for idx, row in enumerate(table_data):
        if not isinstance(row, dict):
            continue
        val = row.get('col_level')
        if val is not None:
            try:
                if '折扣' in str(val):
                    discount_rows.append(row.get('_row_num', idx + 1))
            except Exception:
                continue

    if discount_rows:
        row_list = ', '.join(str(r) for r in discount_rows)
        msg = f"[{table_id_str or 'Table'}] Warning: '折扣' (discount) detected in col_level on {len(discount_rows)} row(s): row(s) {row_list}."
        logging.warning(msg)
        if monitor:
            monitor.log_warning(msg)


def _is_valid_cell_value(val: Any, is_numeric: bool = False) -> bool:
    """Checks if a cell value is present and non-zero (if numeric)."""
    if val is None:
        return False
    val_str = str(val).strip()
    if not val_str:
        return False
    if not is_numeric:
        return True
    try:
        return decimal.Decimal(val_str.replace(',', '')) != 0
    except (decimal.InvalidOperation, ValueError, TypeError):
        try:
            from .data_processor.cbm import _calculate_single_cbm
            cbm_val = _calculate_single_cbm(val_str, 0)
            return cbm_val is not None and cbm_val != 0
        except Exception:
            return False


def validate_table_data_presence(
    current_table_data: List[Dict[str, Any]], 
    table_id_str: str, 
    column_mapping: Dict[str, str], 
    monitor: Optional[Any] = None
):
    """
    Validates that required columns exist in headers and contain valid values across rows.
    """
    validate_col_level_discount(current_table_data, table_id_str, monitor=monitor)

    ALWAYS_REQUIRED = ['col_po', 'col_item', 'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm']
    PRICING_COLUMNS = ['col_amount', 'col_unit_price', 'col_qty_sf']

    if not current_table_data:
        logging.warning(f"Validation: {table_id_str} has no data rows to validate.")
        return

    first_row = current_table_data[0]
    required_cols = list(ALWAYS_REQUIRED) + [c for c in PRICING_COLUMNS if c in column_mapping]
    missing_data_cols = [f"{c} (Missing Header)" for c in required_cols if c not in column_mapping]

    for col in required_cols:
        if col in column_mapping:
            is_num = col not in ('col_po', 'col_item')
            if not _is_valid_cell_value(first_row.get(col), is_numeric=is_num):
                missing_data_cols.append(col)

    if missing_data_cols:
        row_num = first_row.get('_row_num', '?')
        err_msg = f"Row {row_num}: Missing columns: {', '.join(missing_data_cols)}"
        if monitor:
            monitor.log_process_item(f"{table_id_str} First-Row Validation", status="error", error=err_msg)
        raise DataValidationError(err_msg)

    # Validate all rows for critical identifiers, qty & pricing columns
    STRICT_ALL_ROWS_COLS = ['col_po', 'col_item', 'col_qty_pcs', 'col_qty_sf', 'col_unit_price', 'col_amount']
    check_all_cols = [c for c in STRICT_ALL_ROWS_COLS if c in column_mapping]

    for row in current_table_data:
        missing_in_row = [
            c for c in check_all_cols
            if not _is_valid_cell_value(row.get(c), is_numeric=(c not in ('col_po', 'col_item')))
        ]
        if missing_in_row:
            row_num = row.get('_row_num', '?')
            err_msg = f"Row {row_num}: Missing or zero value for: {', '.join(missing_in_row)}"
            if monitor:
                monitor.log_process_item(f"{table_id_str} Row Validation", status="error", error=err_msg)
            raise DataValidationError(err_msg)

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
    prefix = "[validate_weight_integrity]"
    table_context = f" in {table_id_str}" if table_id_str else ""
    
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

        try:
            # Handle Decimals, floats, or strings
            net_val = net_raw if isinstance(net_raw, decimal.Decimal) else DataConverter.convert_to_decimal(net_raw)
            gross_val = gross_raw if isinstance(gross_raw, decimal.Decimal) else DataConverter.convert_to_decimal(gross_raw)
            
            # Skip if both are missing (legitimate filler or spacer row)
            if net_val is None and gross_val is None:
                continue

            # --- NEW STRICT VALIDATION: Unpaired weights are forbidden ---
            # If one exists but the other doesn't, it's a data entry error
            if net_val is None or gross_val is None:
                row_num = row.get('_row_num', '?')
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
                row_num = row.get('_row_num', '?')
                error_msg = f"Row {row_num}: Gross <= Net (Net = {net_val:.2f}, Gross = {gross_val:.2f})"
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
                    row_num = row.get('_row_num', '?')
                    error_msg = f"Row {row_num}: Tare mismatch (expected Gross = {expected_gross:.2f}, but found {gross_val:.2f})"
                    if not ignore_tare_warning:
                        raise DataValidationError(error_msg)
                    else:
                        logging.warning(f"{prefix} Ignored: {error_msg}")
                        if monitor:
                            monitor.log_warning(f"Ignored: {error_msg}")

        except (decimal.InvalidOperation, ValueError, TypeError):
            continue

def validate_cbm_pcs_proportion(
    data_rows: List[Dict[str, Any]], 
    table_id_str: str, 
    column_mapping: Dict[str, str], 
    monitor: Optional[Any] = None,
    ignore_cbm_warning: bool = False
):
    """
    Validates CBM and PCS (or other basis column) proportions to detect:
    - Zero CBM when basis (PCS/SF) > 0.
    - Abnormally high ratio (> 0.5 CBM/unit).
    - Monotonicity violations within each PO (A having more pieces than B, but A having less CBM).
    - Ratio outliers (> 10x or < 0.1x of the median ratio across the table, if >= 3 rows have valid ratios).
    """
    prefix = "[validate_cbm_pcs_proportion]"
    
    # 1. Identify canonical columns
    col_po = 'col_po'
    col_cbm = 'col_cbm'
    
    # Determine the basis column (prefer col_qty_pcs if mapped, then col_qty_sf)
    basis_col = 'col_qty_pcs'
    if basis_col not in column_mapping and 'col_qty_sf' in column_mapping:
        basis_col = 'col_qty_sf'

    # Skip if CBM or the basis column are not mapped
    if col_cbm not in column_mapping:
        logging.debug(f"{prefix} CBM column ({col_cbm}) not in column_mapping, skipping CBM validation.")
        return
    if basis_col not in column_mapping:
        logging.debug(f"{prefix} Basis column ({basis_col}) not in column_mapping, skipping CBM validation.")
        return

    # Skip if the table contains no non-zero CBM values
    has_any_cbm = False
    for row in data_rows:
        cbm_val = row.get(col_cbm)
        if cbm_val is not None:
            try:
                cbm_dec = cbm_val if isinstance(cbm_val, decimal.Decimal) else DataConverter.convert_to_decimal(cbm_val)
                if cbm_dec is not None and cbm_dec > 0:
                    has_any_cbm = True
                    break
            except (decimal.InvalidOperation, ValueError, TypeError):
                continue
    if not has_any_cbm:
        logging.debug(f"{prefix} Table has no positive CBM values, skipping proportion validation.")
        return

    # Collect ALL violations instead of failing on the first one
    violations = []

    # 2. Check each row for Zero CBM and Abnormally High Ratio
    for row in data_rows:
        basis_raw = row.get(basis_col)
        cbm_raw = row.get(col_cbm)
        po_val = row.get(col_po, "Unknown PO")
        item_val = row.get('col_item', "Unknown Item")
        
        basis_dec = None
        if basis_raw is not None:
            try:
                basis_dec = basis_raw if isinstance(basis_raw, decimal.Decimal) else DataConverter.convert_to_decimal(basis_raw)
            except (decimal.InvalidOperation, ValueError, TypeError):
                pass
        
        if basis_dec is None or basis_dec <= 0:
            continue

        cbm_dec = None
        if cbm_raw is not None:
            try:
                cbm_dec = cbm_raw if isinstance(cbm_raw, decimal.Decimal) else DataConverter.convert_to_decimal(cbm_raw)
            except (decimal.InvalidOperation, ValueError, TypeError):
                pass

        # Check for Zero CBM
        if cbm_dec is None or cbm_dec == 0:
            row_num = row.get('_row_num', '?')
            violations.append(
                f"Row {row_num}: Zero CBM ({basis_dec:.0f}pcs)"
            )
            continue  # Skip ratio check if CBM is zero

        # Check for Abnormally High Ratio (> 0.5 CBM/unit)
        ratio = cbm_dec / basis_dec
        if ratio > decimal.Decimal('0.5'):
            row_num = row.get('_row_num', '?')
            violations.append(
                f"Row {row_num}: High ratio ({ratio:.2f} > 0.5)"
            )

    # 3. Monotonicity Check within PO Group
    po_groups = {}
    for row in data_rows:
        po_val = row.get(col_po)
        if not po_val:
            continue
        po_str = str(po_val).strip()
        if not po_str:
            continue
        basis_raw = row.get(basis_col)
        cbm_raw = row.get(col_cbm)
        try:
            basis_dec = basis_raw if isinstance(basis_raw, decimal.Decimal) else DataConverter.convert_to_decimal(basis_raw)
            cbm_dec = cbm_raw if isinstance(cbm_raw, decimal.Decimal) else DataConverter.convert_to_decimal(cbm_raw)
            if basis_dec is not None and basis_dec > 0:
                if cbm_dec is None:
                    cbm_dec = decimal.Decimal(0)
                if po_str not in po_groups:
                    po_groups[po_str] = []
                po_groups[po_str].append({
                    'row': row,
                    'basis': basis_dec,
                    'cbm': cbm_dec,
                    'item': row.get('col_item', 'Unknown Item')
                })
        except (decimal.InvalidOperation, ValueError, TypeError):
            continue

    tolerance = decimal.Decimal('0.0001')
    cbm_tol_pct = decimal.Decimal('0.032') # 3.2% relative tolerance
    for po, group_rows in po_groups.items():
        n = len(group_rows)
        if n < 2:
            continue
            
        # O(N) ratio validation against group median ratio
        # Sort group by basis ascending so adjacent comparisons suffice (O(N log N) / O(N))
        sorted_group = sorted(group_rows, key=lambda x: x['basis'])
        
        # Calculate group median ratio
        group_ratios = [r['cbm'] / r['basis'] for r in sorted_group if r['basis'] > 0 and r['cbm'] > 0]
        if group_ratios:
            sorted_ratios = sorted(group_ratios)
            m_len = len(sorted_ratios)
            if m_len % 2 == 1:
                group_median_ratio = sorted_ratios[m_len // 2]
            else:
                group_median_ratio = (sorted_ratios[m_len // 2 - 1] + sorted_ratios[m_len // 2]) / decimal.Decimal(2)
        else:
            group_median_ratio = None

        # Perform adjacent linear checks O(N) on sorted group
        for i in range(n - 1):
            row_low = sorted_group[i]       # smaller or equal basis
            row_high = sorted_group[i + 1]   # larger or equal basis
            
            if row_high['basis'] > row_low['basis']:
                allowed_cbm = row_low['cbm'] * (decimal.Decimal('1.0') - cbm_tol_pct)
                if row_high['cbm'] < allowed_cbm - tolerance:
                    row_A_num = row_low['row'].get('_row_num', '?')
                    row_B_num = row_high['row'].get('_row_num', '?')
                    if isinstance(row_A_num, int) and isinstance(row_B_num, int) and row_A_num > row_B_num:
                        row_A_num, row_B_num = row_B_num, row_A_num
                    discrepancy = row_low['cbm'] - row_high['cbm']
                    violations.append(
                        f"Row {row_A_num} & {row_B_num}: Qty/CBM wrong ({row_high['basis']:.0f}pcs={row_high['cbm']:.2f}, {row_low['basis']:.0f}pcs={row_low['cbm']:.2f}, discrepancy {discrepancy:.2f} CBM)"
                    )

    # 4. Check for Ratio Outliers
    ratios = []
    ratio_details = []
    for row in data_rows:
        basis_raw = row.get(basis_col)
        cbm_raw = row.get(col_cbm)
        try:
            basis_dec = basis_raw if isinstance(basis_raw, decimal.Decimal) else DataConverter.convert_to_decimal(basis_raw)
            cbm_dec = cbm_raw if isinstance(cbm_raw, decimal.Decimal) else DataConverter.convert_to_decimal(cbm_raw)
            if basis_dec is not None and basis_dec > 0 and cbm_dec is not None and cbm_dec > 0:
                ratio = cbm_dec / basis_dec
                ratios.append(ratio)
                ratio_details.append({
                    'row': row,
                    'ratio': ratio,
                    'cbm': cbm_dec,
                    'basis': basis_dec,
                    'po': row.get(col_po, 'Unknown PO'),
                    'item': row.get('col_item', 'Unknown Item')
                })
        except (decimal.InvalidOperation, ValueError, TypeError):
            continue
    
    if len(ratios) >= 3:
        sorted_ratios = sorted(ratios)
        n = len(sorted_ratios)
        if n % 2 == 1:
            median_ratio = sorted_ratios[n // 2]
        else:
            median_ratio = (sorted_ratios[n // 2 - 1] + sorted_ratios[n // 2]) / decimal.Decimal(2)
            
        upper_limit = median_ratio * decimal.Decimal(10)
        lower_limit = median_ratio * decimal.Decimal('0.1')
        
        for rd in ratio_details:
            row_num = rd['row'].get('_row_num', '?')
            if rd['ratio'] > upper_limit or rd['ratio'] < lower_limit:
                violations.append(
                    f"Row {row_num}: Outlier ({rd['ratio']:.2f} vs median {median_ratio:.2f})"
                )

    # --- Report all violations at once ---
    if violations:
        summary = f"CBM Errors:\n" + "\n".join(violations)
        
        if ignore_cbm_warning:
            logging.warning(f"[{table_id_str}] [CBM Validation Bypassed]: {summary}")
            if monitor:
                monitor.log_warning(f"Ignored CBM Warnings: {summary}")
        else:
            if monitor:
                monitor.log_process_item(f"{table_id_str} CBM Validation", status="error", error=summary)
            raise DataValidationError(summary)

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
    import re
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
            pallet_id = str(raw_id).strip() if raw_id is not None else ""
            
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
            
            # Determine 1/0 boundary marker count
            try:
                count = 1 if (raw_count is not None and int(float(str(raw_count).strip())) >= 1) else 0
            except (ValueError, TypeError):
                count = 0
                
            raw_id = row.get(pallet_id_key)
            pallet_id = str(raw_id).strip() if raw_id is not None else ""
            
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
            try:
                count = 1 if (raw_count is not None and int(float(str(raw_count).strip())) >= 1) else 0
            except (ValueError, TypeError):
                count = 0

            raw_id = row.get(pallet_id_key)
            pallet_id = str(raw_id).strip() if raw_id is not None else ""
            pallet_label = f" (Pallet '{pallet_id}')" if pallet_id else ""

            if count == 0:
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


def validate_data(
    data_rows: List[Dict[str, Any]], 
    table_id_str: str, 
    column_mapping: Dict[str, str], 
    monitor: Optional[Any] = None,
    phase: str = 'presence',
    ignore_tare_warning: bool = False,
    ignore_cbm_warning: bool = False
):
    """
    Unified entry point for all table-level data validation.
    
    Args:
        data_rows: The list of row dictionaries to validate.
        table_id_str: Human readable ID (e.g. "Table 1") for error messages.
        column_mapping: The canonical-to-Excel-letter mapping.
        monitor: Optional PipelineMonitor for logging errors.
        phase: 'presence' (check for at least one value per required col) 
               or 'integrity' (check weight integrity and tare matching)
               or 'cbm_proportion' (check CBM and PCS proportions).
    """
    if phase == 'presence':
        validate_table_data_presence(data_rows, table_id_str, column_mapping, monitor=monitor)
    elif phase == 'integrity':
        validate_weight_integrity(
            data_rows, 
            table_id_str=table_id_str, 
            monitor=monitor, 
            ignore_tare_warning=ignore_tare_warning
        )
        verify_pallet_integrity(
            data_rows,
            table_id_str=table_id_str,
            monitor=monitor
        )
    elif phase == 'cbm_proportion':
        validate_cbm_pcs_proportion(
            data_rows, 
            table_id_str, 
            column_mapping, 
            monitor=monitor, 
            ignore_cbm_warning=ignore_cbm_warning
        )
    else:
        raise ValueError(f"Unknown validation phase requested: {phase}. Supported: 'presence', 'integrity', 'cbm_proportion'.")


logging.info("[validation] Module loaded with consolidated validation routines.")
