import re
from typing import Any, Dict, Optional, Union
from .models import MappingContext


def apply_fallback(
    row_dict: Dict[str, Any],
    target_id: str,
    mapping_rule: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool
):
    """Applies fallback value based on execution flags."""
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
    DAF_mode: bool,
    custom_mode: bool
) -> Optional[Union[str, Dict[str, Any]]]:
    """Resolves formula for current mode."""
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


class RuleEngine:
    """Stateless evaluator for column mapping rules."""

    @staticmethod
    def evaluate_column_rule(
        target_id: str,
        mapping_rule: Dict[str, Any],
        row_dict: Dict[str, Any],
        context: MappingContext
    ) -> None:
        """Applies mapping rules (fallback, static values, formula directives) in-place."""
        if not isinstance(mapping_rule, dict):
            return

        if "static_value" in mapping_rule:
            row_dict[target_id] = mapping_rule["static_value"]
            return

        apply_fallback(
            row_dict=row_dict,
            target_id=target_id,
            mapping_rule=mapping_rule,
            DAF_mode=context.DAF_mode,
            custom_mode=context.custom_mode
        )

        formula = resolve_mode_formula(mapping_rule, DAF_mode=context.DAF_mode, custom_mode=context.custom_mode)
        if formula is not None:
            parsed = parse_formula_def(formula)
            if parsed:
                row_dict[f"{target_id}_formula"] = parsed


def parse_mapping_rules(
    mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str]
) -> Dict[str, Any]:
    """Parses mapping rules into static value maps, dynamic rules, and formulas."""
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

    for rule_key, rule_value in mapping_rules.items():
        if not isinstance(rule_value, dict):
            continue

        if rule_key == "data_map":
            parsed_result["dynamic_mapping_rules"].update(rule_value)
            continue

        rule_type = rule_value.get("type")

        if rule_type == "initial_static_rows":
            static_column_id = rule_value.get("column_header_id")
            target_col_idx = column_id_map.get(static_column_id)

            if target_col_idx is not None:
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

        if rule_type == "formula":
            parsed_result["formula_rules"][target_id] = {
                "template": rule_value.get("formula_template"),
                "input_ids": rule_value.get("inputs", [])
            }
        elif "static_value" in rule_value:
            parsed_result["static_value_map"][target_id] = rule_value["static_value"]
        else:
            parsed_result["dynamic_mapping_rules"][rule_key] = rule_value

    for col_id in column_id_map:
        if col_id not in covered_col_ids and col_id != "col_static":
            parsed_result["dynamic_mapping_rules"][col_id] = {"column": col_id}

    return parsed_result


__all__ = [
    "RuleEngine",
    "apply_fallback",
    "parse_formula_def",
    "resolve_mode_formula",
    "parse_mapping_rules"
]
