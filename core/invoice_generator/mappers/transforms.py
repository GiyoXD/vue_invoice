import logging
from typing import Any, Dict, Iterable, List, Optional, Tuple, Union
from .models import InvoiceMode
from .rules import (
    apply_fallback,
    get_fallback_value,
    parse_formula_def,
    parse_mapping_rules,
    resolve_mode_formula,
)
from .source_normalizer import normalize_source_to_rows, extract_cell_value

logger = logging.getLogger(__name__)




def _resolve_formula_for_rule(
    rule: Dict[str, Any],
    inv_mode: InvoiceMode
) -> Optional[Dict[str, Any]]:
    mode_formula = resolve_mode_formula(rule, invoice_format=inv_mode.invoice_format)
    if not mode_formula:
        return None

    if inv_mode.pricing_net_weight:
        if isinstance(mode_formula, str):
            mode_formula = mode_formula.replace('{col_qty_sf}', '{col_net}')
        elif isinstance(mode_formula, dict):
            mode_formula = dict(mode_formula)
            # If template has raw {col_qty_sf}, replace it
            if 'template' in mode_formula and isinstance(mode_formula['template'], str):
                mode_formula['template'] = mode_formula['template'].replace('{col_qty_sf}', '{col_net}')
            # If inputs list exists, update input token
            if 'inputs' in mode_formula and isinstance(mode_formula['inputs'], list):
                mode_formula['inputs'] = [
                    'col_net' if inp == 'col_qty_sf' else inp
                    for inp in mode_formula['inputs']
                ]

    parsed_formula = parse_formula_def(mode_formula)
    if parsed_formula:
        inputs = parsed_formula.get('inputs', [])
        return {
            'type': 'formula',
            'template': parsed_formula['template'],
            'inputs': inputs
        }
    return None


def _resolve_fallback_for_rule(
    row_dict: Dict[str, Any],
    source_key: str,
    target_id: str,
    rule: Dict[str, Any],
    row: Dict[str, Any],
    dynamic_mapping_rules: Dict[str, Any],
    inv_mode: InvoiceMode
) -> None:
    if row_dict.get(target_id) in [None, ""]:
        if source_key == 'col_desc':
            po_rule = dynamic_mapping_rules.get('col_po')
            po_val = extract_cell_value(row, 'col_po', po_rule)
            if po_val in [None, ""]:
                return
        apply_fallback(row_dict, target_id, rule, invoice_format=inv_mode.invoice_format)


def _build_row_dict(
    row: Dict[str, Any],
    dynamic_mapping_rules: Optional[Dict[str, Any]] = None,
    static_value_map: Optional[Dict[str, Any]] = None,
    inv_mode: Optional[InvoiceMode] = None
) -> Dict[str, Any]:
    if inv_mode is None:
        inv_mode = InvoiceMode()
    dynamic_mapping_rules = dynamic_mapping_rules or {}
    static_value_map = static_value_map or {}
    row_dict: Dict[str, Any] = {}

    for source_key, rule in dynamic_mapping_rules.items():
        if not isinstance(rule, dict) or not source_key:
            continue

        target_id = source_key
        val = extract_cell_value(row, source_key, rule)
        if val is not None:
            row_dict[target_id] = val

        formula = _resolve_formula_for_rule(rule, inv_mode)
        if formula:
            row_dict[target_id] = formula
            continue

        _resolve_fallback_for_rule(
            row_dict, source_key, target_id, rule, row, dynamic_mapping_rules, inv_mode
        )

    for col_id, static_val in static_value_map.items():
        row_dict[col_id] = static_val

    return row_dict


def prepare_data_rows(
    data_source: Any,
    dynamic_mapping_rules: Optional[Dict[str, Any]] = None,
    static_value_map: Optional[Dict[str, Any]] = None,
    mode: Optional[Union[InvoiceMode, str, Dict[str, Any]]] = None,
) -> Tuple[List[Dict[str, Any]], int]:
    if isinstance(mode, InvoiceMode):
        inv_mode = mode
    elif isinstance(mode, str):
        inv_mode = InvoiceMode(invoice_format=mode)
    else:
        inv_mode = InvoiceMode.from_dict(mode)
    normalized_rows = normalize_source_to_rows(data_source)
    data_rows = []
    for row in normalized_rows:
        row_dict = _build_row_dict(
            row=row,
            dynamic_mapping_rules=dynamic_mapping_rules,
            static_value_map=static_value_map,
            inv_mode=inv_mode
        )
        data_rows.append(row_dict)

    return data_rows, len(data_rows)


def resolve_static_placeholders(
    static_payload: Dict[str, Any],
    desc_fallback_str: str = ""
) -> Dict[str, Any]:
    """Resolve placeholders in static_payload dictionary and return resolved static_payload."""
    if not static_payload or 'col_static' not in static_payload:
        return static_payload or {}

    desc_str = desc_fallback_str or static_payload.get('description_fallback', "")
    resolved_static = dict(static_payload)
    static_values = static_payload.get('col_static', [])
    if isinstance(static_values, list):
        new_static_values = []
        for val in static_values:
            if isinstance(val, str) and "{col_desc_fallback}" in val:
                val = val.replace("{col_desc_fallback}", str(desc_str or ""))
            new_static_values.append(val)
        resolved_static['col_static'] = new_static_values

    return resolved_static


def populate_static_content(
    data_rows: List[Dict[str, Any]],
    static_payload: Dict[str, Any]
) -> None:
    """Populate resolved static payload into data rows."""
    if data_rows is None or not static_payload or 'col_static' not in static_payload:
        return

    static_values = static_payload.get('col_static')
    if not static_values:
        return

    static_col_id = 'col_static'
    
    if static_values:
        num_static_values = len(static_values)
        while len(data_rows) < num_static_values:
            data_rows.append({})

        for i, static_value in enumerate(static_values):
            data_rows[i][static_col_id] = static_value


def format_pallet_counts(
    data_rows: List[Dict[str, Any]],
    num_data_rows: int,
    pallet_col_id: Optional[str]
) -> None:
    """Carries non-empty values forward across rows for vertical cell merging."""
    if not data_rows or num_data_rows <= 0 or not pallet_col_id:
        return

    carry_value = None
    for row in data_rows[:num_data_rows]:
        raw_val = row.get(pallet_col_id)
        if raw_val not in (None, "", 0):
            carry_value = raw_val
        elif carry_value is not None:
            row[pallet_col_id] = carry_value


def transform_table_data(
    data_source: Any,
    mapping_rules: Dict[str, Any],
    static_payload: Optional[Dict[str, Any]] = None,
    mode: Optional[Union[InvoiceMode, str, Dict[str, Any]]] = None,
    target_columns: Optional[Iterable[str]] = None
) -> Tuple[List[Dict[str, Any]], int]:
    """Main entry point for table data transformation pipeline."""
    mapping_rules = mapping_rules or {}
    static_payload = static_payload or {}

    if isinstance(mode, InvoiceMode):
        inv_mode = mode
    elif isinstance(mode, str):
        inv_mode = InvoiceMode(invoice_format=mode)
    else:
        inv_mode = InvoiceMode.from_dict(mode)

    target_cols = {target_columns} if isinstance(target_columns, str) else (set(target_columns) if target_columns else set())

    parsed = parse_mapping_rules(
        mapping_rules=mapping_rules,
        target_columns=target_cols
    )

    desc_rule = parsed['dynamic_mapping_rules'].get('col_desc', {})
    desc_fallback_str = get_fallback_value(desc_rule, invoice_format=inv_mode.invoice_format) or ""

    resolved_static = resolve_static_placeholders(
        static_payload=static_payload,
        desc_fallback_str=str(desc_fallback_str)
    )

    data_rows, num_data_rows = prepare_data_rows(
        data_source=data_source,
        dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
        static_value_map=parsed['static_value_map'],
        mode=inv_mode
    )

    populate_static_content(
        data_rows=data_rows,
        static_payload=resolved_static
    )

    pallet_col_id = 'col_pallet_count' if ('col_pallet_count' in target_cols or (not target_cols and 'col_pallet_count' in parsed['dynamic_mapping_rules'])) else None
    if pallet_col_id:
        format_pallet_counts(
            data_rows=data_rows,
            num_data_rows=num_data_rows,
            pallet_col_id=pallet_col_id
        )

    return data_rows, num_data_rows


__all__ = [
    "prepare_data_rows",
    "resolve_static_placeholders",
    "populate_static_content",
    "format_pallet_counts",
    "transform_table_data",
]


