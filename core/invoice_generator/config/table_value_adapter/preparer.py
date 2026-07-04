from typing import Any, Union, Dict, List, Tuple, Optional
from decimal import Decimal
import logging
import re
logger = logging.getLogger(__name__)


def parse_mapping_rules(
    mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str]
) -> Dict[str, Any]:
    """
    Parses the mapping rules from a standardized, ID-based configuration.
    """
    # --- Initialize all return values ---
    parsed_result = {
        "static_value_map": {},
        "initial_static_col1_values": [],
        "dynamic_mapping_rules": {},
        "formula_rules": {},
        "col1_index": -1,
        "num_static_labels": 0,
        "static_column_header_name": None,
        "apply_special_border_rule": False
    }

    covered_col_ids = set()

    # --- Process all rules in a single, intelligent pass ---
    for rule_key, rule_value in mapping_rules.items():
        if not isinstance(rule_value, dict):
            continue

        if rule_key == "data_map":
            parsed_result["dynamic_mapping_rules"].update(rule_value)
            continue

        rule_type = rule_value.get("type")

        # --- Handler for Initial Static Rows ---
        if rule_type == "initial_static_rows":
            static_column_id = rule_value.get("column_header_id")
            target_col_idx = column_id_map.get(static_column_id)

            if target_col_idx:
                parsed_result["static_column_header_name"] = idx_to_header_map.get(target_col_idx)
                parsed_result["col1_index"] = target_col_idx
                parsed_result["initial_static_col1_values"] = rule_value.get("values", [])
                parsed_result["num_static_labels"] = len(parsed_result["initial_static_col1_values"])
                
                parsed_result["formula_rules"][static_column_id] = {
                    "template": rule_value.get("formula_template"),
                    "input_ids": rule_value.get("inputs", [])
                }
            continue

        target_id = rule_key
        if target_id:
            covered_col_ids.add(target_id)

        # --- Handler for Formulas ---
        if rule_type == "formula":
            parsed_result["formula_rules"][target_id] = {
                "template": rule_value.get("formula_template"),
                "input_ids": rule_value.get("inputs", [])
            }

        # --- Handler for Static Values ---
        elif "static_value" in rule_value:
            parsed_result["static_value_map"][target_id] = rule_value["static_value"]
        
        # --- Handler for top-level Dynamic Rules (used by 'aggregation') ---
        else:
            parsed_result["dynamic_mapping_rules"][rule_key] = rule_value
            
    # --- Auto-Mapping: Add default rules for any column ID not explicitly covered ---
    for col_id in column_id_map:
        if col_id not in covered_col_ids and col_id != "col_static":
            parsed_result["dynamic_mapping_rules"][col_id] = {"column": col_id}

    return parsed_result


def _to_numeric(value: Any) -> Union[int, float, None, Any]:
    """
    Safely attempts to convert a value to a float or int.
    """
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        try:
            cleaned_val = value.replace(',', '').strip()
            if not cleaned_val:
                return None
            return float(cleaned_val) if '.' in cleaned_val else int(cleaned_val)
        except (ValueError, TypeError):
            return value
    if isinstance(value, Decimal):
        if value == value.to_integral_value():
            return int(value)
        return float(str(value))
    return value


def _apply_fallback(
    row_dict: Dict[str, Any],
    target_id: str,
    mapping_rule: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool
):
    """
    Applies a fallback value to the row_dict based on the DAF_mode and custom_mode.
    """
    fallback_config = mapping_rule.get('fallback')
    if isinstance(fallback_config, dict):
        if DAF_mode and 'daf' in fallback_config:
            row_dict[target_id] = fallback_config['daf']
            return
        elif custom_mode and 'custom' in fallback_config:
            row_dict[target_id] = fallback_config['custom']
            return
        elif 'standard' in fallback_config:
            row_dict[target_id] = fallback_config['standard']
            return
        elif 'default' in fallback_config:
            row_dict[target_id] = fallback_config['default']
            return
    elif fallback_config is not None:
        row_dict[target_id] = fallback_config
        return


def _parse_formula_def(formula_def: Union[str, Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """
    Normalizes a formula definition into a dictionary with 'template' and 'inputs'.
    """
    if isinstance(formula_def, dict) and 'template' in formula_def:
        return formula_def
        
    if isinstance(formula_def, str) and formula_def.strip():
        inputs = re.findall(r'\{([^}]+)\}', formula_def)
        filtered_inputs = [inp for inp in inputs if inp != 'row']
        
        processed_template = formula_def
        for i, input_col in enumerate(filtered_inputs):
            processed_template = processed_template.replace(f'{{{input_col}}}', f'{{col_ref_{i}}}')
            
        return {
            'template': processed_template,
            'inputs': filtered_inputs
        }
        
    return None


def _resolve_mode_formula(
    rule: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool
) -> Optional[Union[str, Dict[str, Any]]]:
    """
    Resolves the correct formula from a mapping rule based on the current mode.
    """
    formula_config = rule.get('formula')
    if formula_config is None:
        return None

    if isinstance(formula_config, dict):
        if custom_mode and 'custom' in formula_config:
            return formula_config['custom']
        elif DAF_mode and 'daf' in formula_config:
            return formula_config['daf']
        elif not custom_mode and not DAF_mode and 'standard' in formula_config:
            return formula_config['standard']
        return None

    if isinstance(formula_config, str):
        return formula_config

    return None


def _get_value_from_source(source_container: Any, rule_key: str, row_idx: int = None) -> Any:
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

        mode_formula = _resolve_mode_formula(rule, DAF_mode, custom_mode)
        if mode_formula:
            if pricing_net_weight and re.search(r'\{\s*col_qty_sf\s*\}', mode_formula):
                mode_formula = re.sub(r'\{\s*col_qty_sf\s*\}', '{col_net}', mode_formula)
                
            parsed_formula = _parse_formula_def(mode_formula)
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
            _apply_fallback(row_dict, target_id, rule, DAF_mode, custom_mode)

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