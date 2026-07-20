import logging
import re
from typing import Any, Dict, List, Optional, Tuple, Union

from ..rules import apply_fallback, parse_formula_def, resolve_mode_formula

logger = logging.getLogger(__name__)


def _get_value_from_source(source_container: Any, rule_key: str, row_idx: Optional[int] = None) -> Any:
    """
    Extract a value from the data source using the rule key.
    """
    if rule_key is None or not isinstance(source_container, dict):
        return None

    if row_idx is not None:
        col_data = source_container.get(rule_key)
        if isinstance(col_data, list) and row_idx < len(col_data):
            return col_data[row_idx]

    elif rule_key in source_container:
        return source_container[rule_key]

    return None


def _build_row_dict(
    source_container: Any,
    row_idx: Optional[int],
    dynamic_mapping_rules: Dict[str, Any],
    parent_column_ids: List[str],
    static_value_map: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool,
    pricing_net_weight: bool = False
) -> Dict[str, Any]:
    """
    Build a single row dictionary by applying mapping rules, formulas, fallbacks, and static values.
    """
    row_dict = {}

    for source_key, rule in dynamic_mapping_rules.items():
        if not isinstance(rule, dict):
            continue

        target_id = source_key
        if not target_id:
            continue
        if target_id in parent_column_ids:
            continue

        val = _get_value_from_source(source_container, source_key, row_idx)
        if val is not None:
            row_dict[target_id] = val

        mode_formula = resolve_mode_formula(rule, DAF_mode, custom_mode)
        if mode_formula:
            if pricing_net_weight and re.search(r'\{\s*col_qty_sf\s*\}', mode_formula):
                mode_formula = re.sub(r'\{\s*col_qty_sf\s*\}', '{col_net}', mode_formula)
                
            parsed_formula = parse_formula_def(mode_formula)
            if parsed_formula:
                inputs = parsed_formula.get('inputs', [])
                row_dict[target_id] = {
                    'type': 'formula',
                    'template': parsed_formula['template'],
                    'inputs': inputs
                }
                continue

        if row_dict.get(target_id) in [None, ""]:
            if source_key == 'col_desc':
                po_val = _get_value_from_source(source_container, 'col_po', row_idx)
                if not po_val:
                    continue
            apply_fallback(row_dict, target_id, rule, DAF_mode, custom_mode)

    # Apply static values
    for col_id, static_val in static_value_map.items():
        if col_id not in row_dict:
            row_dict[col_id] = static_val

    return row_dict


def prepare_data_rows(
    data_source_type: str,
    data_source: Union[Dict, List],
    dynamic_mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str],
    desc_col_idx: int,
    num_static_labels: int,
    static_value_map: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool = False,
    parent_column_ids: List[str] = None,
    pricing_net_weight: bool = False
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Prepares data rows by applying mapping rules to the data source.
    """
    parent_column_ids = parent_column_ids or []
    data_rows_prepared = []
    num_data_rows_from_source = 0

    build_kwargs = {
        'dynamic_mapping_rules': dynamic_mapping_rules,
        'parent_column_ids': parent_column_ids,
        'static_value_map': static_value_map,
        'DAF_mode': DAF_mode,
        'custom_mode': custom_mode,
        'pricing_net_weight': pricing_net_weight,
    }

    if isinstance(data_source, dict):
        for val in data_source.values():
            if isinstance(val, list):
                num_data_rows_from_source = max(num_data_rows_from_source, len(val))

        for i in range(num_data_rows_from_source):
            row_dict = _build_row_dict(source_container=data_source, row_idx=i, **build_kwargs)
            data_rows_prepared.append(row_dict)

    elif isinstance(data_source, list):
        num_data_rows_from_source = len(data_source)

        for row_data in data_source:
            row_dict = _build_row_dict(source_container=row_data, row_idx=None, **build_kwargs)
            data_rows_prepared.append(row_dict)

    if num_static_labels > len(data_rows_prepared):
        data_rows_prepared.extend([{}] * (num_static_labels - len(data_rows_prepared)))

    return data_rows_prepared, num_data_rows_from_source
