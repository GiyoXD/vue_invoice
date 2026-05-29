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
            val_str = str(val).strip()
            if val_str:
                try:
                    num_val = decimal.Decimal(val_str.replace(',', ''))
                    if num_val != 0:
                        has_val = True
                except (decimal.InvalidOperation, ValueError, TypeError):
                    has_val = True
        
        if not has_val:
            missing_data_cols.append(col_name)

    if missing_data_cols:
        row_num = first_row.get('_row_num')
        row_num_str = f" (Row {row_num})" if row_num else ""
        err_msg = (
            f"Data Validation Error{row_num_str}: {table_id_str} is missing mandatory data in the first row for: "
            f"[{', '.join(missing_data_cols)}]. "
            "Please ensure the first row of every table in your Excel is fully populated."
        )
        if monitor:
            monitor.log_process_item(f"{table_id_str} First-Row Validation", status="error", error=err_msg)
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
                po_val = row.get(po_key, "Unknown PO")
                item_val = row.get(item_key, "Unknown Item")
                missing_col = "Gross Weight" if gross_val is None else "Net Weight"
                present_col = "Net Weight" if net_val is not None else "Gross Weight"
                
                row_num = row.get('_row_num')
                row_num_str = f" (Row {row_num})" if row_num else ""
                error_msg = (
                    f"Weight Integrity Error{row_num_str}{table_context}: Partial weight found at row for PO [{po_val}] / Item [{item_val}]. "
                    f"{present_col} has a value, but {missing_col} is missing. "
                    "Weights must always be provided as a Net/Gross pair."
                )
                logging.error(f"{prefix} {error_msg}")
                raise DataValidationError(error_msg)

            # Skip header/footer rows where BOTH are 0
            if net_val == 0 and gross_val == 0:
                continue

            # Validation 1: Strict Positivity Constraint (Gross > Net)
            if gross_val <= net_val:
                po_val = row.get(po_key, "Unknown PO")
                item_val = row.get(item_key, "Unknown Item")
                row_num = row.get('_row_num')
                row_num_str = f" (Row {row_num})" if row_num else ""
                
                error_msg = (
                    f"Weight Validation Error{row_num_str}{table_context}: At row for PO [{po_val}] / Item [{item_val}], "
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
                    row_num = row.get('_row_num')
                    row_num_str = f" (Row {row_num})" if row_num else ""
                    error_msg = (
                        f"Weight Integrity Error{row_num_str}{table_context}: `Net + Pallet Weight` does not equal `Gross Weight` at {row_id_str}. "
                        f"Based on the first row ({ref_row_info}), the Pallet Weight (Tare) is **{reference_tare}**. "
                        f"Expected Gross Weight for this row is {net_val} + {reference_tare} = **{expected_gross}**, "
                        f"but found **{gross_val}**. "
                        "Please ensure all pallets in your table have identical tare weights."
                    )
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
            row_num = row.get('_row_num')
            row_num_str = f" (Row {row_num})" if row_num else ""
            violations.append(
                f"[Zero CBM]{row_num_str}: PO [{po_val}] / Item [{item_val}] has pieces ({basis_dec}) but received 0 CBM. "
                "Missing CBM or incorrect distribution anchor."
            )
            continue  # Skip ratio check if CBM is zero

        # Check for Abnormally High Ratio (> 0.5 CBM/unit)
        ratio = cbm_dec / basis_dec
        if ratio > decimal.Decimal('0.5'):
            row_num = row.get('_row_num')
            row_num_str = f" (Row {row_num})" if row_num else ""
            violations.append(
                f"[High Ratio]{row_num_str}: PO [{po_val}] / Item [{item_val}] has {ratio:.4f} CBM/unit "
                f"({cbm_dec} CBM for {basis_dec} units). Exceeds 0.5 CBM/unit threshold."
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
        for i in range(n):
            for j in range(i + 1, n):
                row_A = group_rows[i]
                row_B = group_rows[j]
                
                if row_A['basis'] > row_B['basis']:
                    # Row A has more pieces, so it should have more CBM.
                    # We allow a tolerance of up to 3.2% of B's CBM.
                    allowed_cbm = row_B['cbm'] * (decimal.Decimal('1.0') - cbm_tol_pct)
                    if row_A['cbm'] < allowed_cbm - tolerance:
                        row_A_num = row_A['row'].get('_row_num')
                        row_B_num = row_B['row'].get('_row_num')
                        rows_str = f" (Rows {row_A_num} and {row_B_num})" if (row_A_num and row_B_num) else f" (Row {row_A_num})" if row_A_num else f" (Row {row_B_num})" if row_B_num else ""
                        violations.append(
                            f"[Monotonicity]{rows_str}: In PO [{po}], Item [{row_A['item']}] has more pieces ({row_A['basis']}) but lower CBM ({row_A['cbm']}) "
                            f"than Item [{row_B['item']}] ({row_B['basis']} pcs, {row_B['cbm']} CBM). Verify pallet count or CBM."
                        )
                elif row_B['basis'] > row_A['basis']:
                    # Row B has more pieces, so it should have more CBM.
                    # We allow a tolerance of up to 3.2% of A's CBM.
                    allowed_cbm = row_A['cbm'] * (decimal.Decimal('1.0') - cbm_tol_pct)
                    if row_B['cbm'] < allowed_cbm - tolerance:
                        row_A_num = row_A['row'].get('_row_num')
                        row_B_num = row_B['row'].get('_row_num')
                        rows_str = f" (Rows {row_A_num} and {row_B_num})" if (row_A_num and row_B_num) else f" (Row {row_A_num})" if row_A_num else f" (Row {row_B_num})" if row_B_num else ""
                        violations.append(
                            f"[Monotonicity]{rows_str}: In PO [{po}], Item [{row_B['item']}] has more pieces ({row_B['basis']}) but lower CBM ({row_B['cbm']}) "
                            f"than Item [{row_A['item']}] ({row_A['basis']} pcs, {row_A['cbm']} CBM). Verify pallet count or CBM."
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
            if rd['ratio'] > upper_limit:
                row_num = rd['row'].get('_row_num')
                row_num_str = f" (Row {row_num})" if row_num else ""
                violations.append(
                    f"[Ratio Outlier]{row_num_str}: PO [{rd['po']}] / Item [{rd['item']}] ratio {rd['ratio']:.6f} is > 10x median ({median_ratio:.6f}). "
                    f"Verify CBM ({rd['cbm']}) or quantity ({rd['basis']})."
                )
            elif rd['ratio'] < lower_limit:
                row_num = rd['row'].get('_row_num')
                row_num_str = f" (Row {row_num})" if row_num else ""
                violations.append(
                    f"[Ratio Outlier]{row_num_str}: PO [{rd['po']}] / Item [{rd['item']}] ratio {rd['ratio']:.6f} is < 0.1x median ({median_ratio:.6f}). "
                    f"Verify CBM ({rd['cbm']}) or quantity ({rd['basis']})."
                )

    # --- Report all violations at once ---
    if violations:
        summary = f"CBM Validation found {len(violations)} issue(s) in {table_id_str}:\n" + "\n".join(f"  {i+1}. {v}" for i, v in enumerate(violations))
        
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
    Validates pallet data integrity in two phases:
    
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
    """
    import re
    pallet_count_key = 'col_pallet_count'
    pallet_id_key = 'col_pallet_id'
    sqft_key = 'col_qty_sf'
    
    has_count = any(pallet_count_key in row for row in table_data)
    has_id = any(pallet_id_key in row for row in table_data)
    
    # If col_pallet_id doesn't exist at all, warn and skip
    if not has_id:
        msg = f"[Pallet Pairing] col_pallet_id column not found in {table_id_str or 'data'}. Skipping pallet validation."
        logging.warning(msg)
        if monitor:
            monitor.log_warning(msg)
        return

    # --- Phase 1: Pairing check (pallet_id must exist for every row with sqft) ---
    missing_id_rows = []
    for idx, row in enumerate(table_data):
        row_num = row.get('_row_num', idx + 1)
        sqft_val = row.get(sqft_key)
        raw_id = row.get(pallet_id_key)
        pallet_id = str(raw_id).strip() if raw_id is not None else ""
        
        # Row has sqft data but no pallet ID → operator forgot to fill
        has_sqft = sqft_val is not None and str(sqft_val).strip() not in ('', '0', '0.0', '0.00')
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

    # --- Phase 2: Correlation check (count vs id) — only if both columns exist ---
    if not has_count:
        return

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
        # We skip validation but reset last_pallet_id so that if a new pallet starts
        # later, we treat it as a new block.
        if not pallet_id and count == 0:
            last_pallet_id = None
            continue

        # Rule C (Presence): If count is 1, pallet_id must be present
        if count == 1 and not pallet_id:
            raise DataValidationError(
                f"Pallet Validation Error (Row {row_num}){f' in {table_id_str}' if table_id_str else ''}: "
                f"Pallet boundary found (count=1), but Pallet ID is missing."
            )

        # Rules A & B: Transitions
        if pallet_id != last_pallet_id:
            # Count MUST be 1 (Rule A)
            if count != 1:
                raise DataValidationError(
                    f"Pallet Validation Error (Row {row_num}){f' in {table_id_str}' if table_id_str else ''}: "
                    f"Pallet ID changed to '{pallet_id}' (from '{last_pallet_id}') "
                    f"but boundary marker count is {raw_count} (expected 1)."
                )
            
            # Rule D: Contiguity (No recurrence after a gap)
            if pallet_id in seen_pallet_ids:
                raise DataValidationError(
                    f"Pallet Validation Error (Row {row_num}){f' in {table_id_str}' if table_id_str else ''}: "
                    f"Pallet ID '{pallet_id}' reappeared after a gap. "
                    f"All rows for a pallet must be contiguous."
                )
            seen_pallet_ids.add(pallet_id)
        else:
            # Continuation of the same pallet ID
            # Count MUST be 0 (Rule B)
            if count != 0:
                raise DataValidationError(
                    f"Pallet Validation Error (Row {row_num}){f' in {table_id_str}' if table_id_str else ''}: "
                    f"Pallet boundary marker count is 1 on row {row_num}, "
                    f"but Pallet ID did not change (still '{pallet_id}')."
                )

        last_pallet_id = pallet_id


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
