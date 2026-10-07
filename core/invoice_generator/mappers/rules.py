import re
from typing import Any, Dict, Iterable, Optional, Union

from .models import VALID_INVOICE_FORMATS


def get_fallback_value(
    mapping_rule: Dict[str, Any],
    invoice_format: str = "standard"
) -> Optional[Any]:
    """Resolves fallback value for a mapping rule (flat scalar only)."""
    if not mapping_rule:
        return None

    fallback = mapping_rule.get('fallback')

    # Only support flat scalar ('fallback' as str, int, float)
    if isinstance(fallback, (str, int, float)):
        val = str(fallback).strip() if isinstance(fallback, str) else fallback
        if val != "":
            return val

    return None


def apply_fallback(
    row_dict: Dict[str, Any],
    target_id: str,
    mapping_rule: Dict[str, Any],
    invoice_format: str = "standard"
):
    """Applies fallback value decisively based on invoice_format."""
    val = get_fallback_value(mapping_rule, invoice_format=invoice_format)
    if val is not None:
        row_dict[target_id] = val


def parse_formula_def(formula_def: Union[str, Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Normalizes formula definition into template and input tokens."""
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


def resolve_mode_formula(
    rule: Dict[str, Any],
    invoice_format: str = "standard"
) -> Optional[Union[str, Dict[str, Any]]]:
    """Resolves formula decisively for the specified invoice_format."""
    fmt = str(invoice_format or "standard").strip().lower()
    if fmt not in VALID_INVOICE_FORMATS:
        raise ValueError(
            f"Invalid invoice_format '{invoice_format}'. "
            f"Must be one of {sorted(VALID_INVOICE_FORMATS)}"
        )
    if not rule:
        return None

    formula_config = rule.get('formula')
    if formula_config is None:
        return None

    if isinstance(formula_config, dict):
        if fmt in formula_config:
            return formula_config[fmt]
        return None

    if isinstance(formula_config, str):
        return formula_config

    return None



def parse_mapping_rules(
    mapping_rules: Dict[str, Any],
    target_columns: Optional[Iterable[str]] = None
) -> Dict[str, Any]:
    """Parses mapping rules into static value maps and dynamic rules."""
    mapping_rules = mapping_rules or {}
    parsed_result = {
        "static_value_map": {},
        "dynamic_mapping_rules": {}
    }

    covered_col_ids = set()

    for rule_key, rule_value in mapping_rules.items():
        if not isinstance(rule_value, dict):
            continue

        if rule_key == "data_map":
            parsed_result["dynamic_mapping_rules"].update(rule_value)
            covered_col_ids.update(rule_value.keys())
            continue

        target_id = rule_key
        if target_id:
            covered_col_ids.add(target_id)

        if "static_value" in rule_value:
            parsed_result["static_value_map"][target_id] = rule_value["static_value"]
        else:
            parsed_result["dynamic_mapping_rules"][rule_key] = rule_value

    target_col_ids = set(target_columns) if target_columns else set()
    if isinstance(target_columns, str):
        target_col_ids = {target_columns}
    for col_id in target_col_ids:
        if col_id not in covered_col_ids and col_id != "col_static":
            parsed_result["dynamic_mapping_rules"][col_id] = {}

    return parsed_result


__all__ = [
    "get_fallback_value",
    "apply_fallback",
    "parse_formula_def",
    "resolve_mode_formula",
    "parse_mapping_rules"
]
