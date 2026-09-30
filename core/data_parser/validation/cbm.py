import logging
import decimal
from typing import List, Dict, Any, Optional
from ..util.converters import DataConverter
from .exceptions import DataValidationError

decimal.getcontext().prec = 28


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
