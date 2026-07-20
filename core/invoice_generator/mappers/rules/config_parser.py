from typing import Any, Dict


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
