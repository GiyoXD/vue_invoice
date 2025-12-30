from typing import Any, Union, Dict, List, Tuple
from decimal import Decimal
import logging
logger = logging.getLogger(__name__)

def parse_mapping_rules(
    mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str]
) -> Dict[str, Any]:
    """
    Parses the mapping rules from a standardized, ID-based configuration.

    This function is refined to handle different mapping structures, such as a
    flat structure for aggregation sheets and a nested 'data_map' for table-based sheets.

    Args:
        mapping_rules: The raw mapping rules dictionary from the sheet's configuration.
        column_id_map: A dictionary mapping column IDs to their 1-based column index.
        idx_to_header_map: A dictionary mapping a column index back to its header text.

    Returns:
        A dictionary containing all the parsed information required for data filling.
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
            continue # Skip non-dictionary rules

        # --- Handler for nested 'data_map' (used by 'processed_tables_multi') ---
        if rule_key == "data_map":
            # The entire dictionary under "data_map" is our set of dynamic rules.
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
                
                parsed_result["formula_rules"][target_col_idx] = {
                    "template": rule_value.get("formula_template"),
                    "input_ids": rule_value.get("inputs", [])
                }
            else:
                logger.warning(f"Warning: Initial static rows column with ID '{static_column_id}' not found.")
            continue

        # For all other rules, get the target column index using the RELIABLE ID
        # Support both legacy 'id' and bundled 'column' keys
        target_id = rule_value.get("id") or rule_value.get("column")
        if target_id:
            covered_col_ids.add(target_id)
        target_col_idx = column_id_map.get(target_id)

        # --- Handler for Formulas ---
        if rule_type == "formula":
            if target_col_idx:
                parsed_result["formula_rules"][target_col_idx] = {
                    "template": rule_value.get("formula_template"),
                    "input_ids": rule_value.get("inputs", [])
                }
            else:
                logger.warning(f"Warning: Could not find target column for formula rule with id '{target_id}'.")

        # --- Handler for Static Values ---
        elif "static_value" in rule_value:
            if target_col_idx:
                parsed_result["static_value_map"][target_col_idx] = rule_value["static_value"]
            else:
                logger.warning(f"Warning: Could not find target column for static_value rule with id '{target_id}'.")
        
        # --- Handler for top-level Dynamic Rules (used by 'aggregation') ---
        else:
            # If it's not a special rule, it's a dynamic mapping rule for the aggregation data type.
            parsed_result["dynamic_mapping_rules"][rule_key] = rule_value
            
    # --- Auto-Mapping: Add default rules for any column ID not explicitly covered ---
    for col_id in column_id_map:
        if col_id not in covered_col_ids and col_id != "col_static":
            # Create a default rule where the key is the col_id itself
            # This enables "Auto-Mapping" where data keys match column IDs
            parsed_result["dynamic_mapping_rules"][col_id] = {"column": col_id}

    return parsed_result

def _to_numeric(value: Any) -> Union[int, float, None, Any]:
    """
    Safely attempts to convert a value to a float or int.
    Handles strings with commas and returns the original value on failure.
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
            return value # Return original string if conversion fails
    if isinstance(value, Decimal):
        return float(value)
    return value # Return original value for other types

def _apply_fallback(
    row_dict: Dict[int, Any],
    target_col_idx: int,
    mapping_rule: Dict[str, Any],
    DAF_mode: bool
):
    """
    Applies a fallback value to the row_dict based on the DAF_mode.
    
    Supports multiple fallback formats:
    1. Bundled config with mode-specific fallbacks:
       "fallback_on_none": "LEATHER", "fallback_on_DAF": "LEATHER"
    2. Bundled config with single fallback (same for both modes):
       "fallback": "LEATHER"
    3. Legacy format (same as #1)
    """
    # Priority 1: Check for mode-specific fallback keys (supports both DAF and non-DAF)
    if DAF_mode:
        if 'fallback_on_DAF' in mapping_rule:
            row_dict[target_col_idx] = mapping_rule['fallback_on_DAF']
            return
    else:
        if 'fallback_on_none' in mapping_rule:
            row_dict[target_col_idx] = mapping_rule['fallback_on_none']
            return
    
    # Priority 2: Try single 'fallback' key (same value for both modes)
    if 'fallback' in mapping_rule:
        row_dict[target_col_idx] = mapping_rule['fallback']
        return
    
    # Priority 3: Fallback to fallback_on_none if nothing else found
    val = mapping_rule.get("fallback_on_none")
    if val is not None:
        row_dict[target_col_idx] = val

def prepare_data_rows(
    data_source_type: str,
    data_source: Union[Dict, List],
    dynamic_mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str],
    desc_col_idx: int,
    num_static_labels: int,
    static_value_map: Dict[int, Any],
    DAF_mode: bool,
) -> Tuple[List[Dict[int, Any]], List[int], bool, int]:
    """
    Prepares data rows by applying mapping rules to the data source.
    Supports both Column-Oriented (Dict of Lists) and Row-Oriented (List of Dicts) sources.
    Uses robust lookup strategies to find data even if keys don't match column IDs exactly.
    """
    
    # Validate description field has fallback - CRITICAL for proper invoice generation
    desc_mapping = None
    for field_name, mapping_rule in dynamic_mapping_rules.items():
        if 'desc' in field_name.lower() and isinstance(mapping_rule, dict):
            desc_mapping = mapping_rule
            break
    
    if desc_mapping:
        has_fallback = any(key in desc_mapping for key in ['fallback_on_none', 'fallback_on_DAF', 'fallback'])
        if not has_fallback:
            logger.warning(f"Description field missing fallback configuration. Recommended: 'fallback_on_none': 'LEATHER'.")
    
    data_rows_prepared = []
    pallet_counts_for_rows = []
    num_data_rows_from_source = 0
    dynamic_desc_used = False
    
    def get_value_from_row_or_cols(source_container: Any, rule: Dict, rule_key: str, row_idx: int = None) -> Any:
        """
        Helper to extract value from source using multiple lookup strategies.
        Args:
            source_container: Either a row dict (row-oriented) or the main data dict (col-oriented)
            rule: The mapping rule dict
            rule_key: The key from the mapping_rules dict (e.g. 'po')
            row_idx: Index of the row (required for column-oriented source)
        """
        possible_keys = []
        
        # 1. Configured source value/key has highest priority
        if 'source_value' in rule: possible_keys.append(rule['source_value'])
        if 'source_key' in rule: possible_keys.append(rule['source_key'])
        
        # 2. Target Column ID (strict mapping)
        target_id = rule.get("column") or rule.get("id")
        if target_id: possible_keys.append(target_id)
        
        # 3. Rule Key (the name in the mappings dict, often 'po', 'item')
        if rule_key: possible_keys.append(rule_key)

        found_value = None
        found = False

        # Try all possible keys
        for key in possible_keys:
            if key is None: continue
            
            # Case A: Column-Oriented (source_container is dict of lists)
            # We need to look up column 'key', then index 'row_idx'
            if row_idx is not None and isinstance(source_container, dict):
                if key in source_container:
                    col_data = source_container[key]
                    if isinstance(col_data, list) and row_idx < len(col_data):
                        found_value = col_data[row_idx]
                        found = True
                        break
            
            # Case B: Row-Oriented (source_container is the row dict/object)
            elif row_idx is None:
                # Direct lookup in the row object
                if isinstance(source_container, dict) and key in source_container:
                    found_value = source_container[key]
                    found = True
                    break
                # Handle tuple keys if the key is an integer index (e.g. source_key: 0)
                elif isinstance(source_container, tuple) and isinstance(key, int):
                     if 0 <= key < len(source_container):
                         found_value = source_container[key]
                         found = True
                         break
        
        if not found:
            return None
        return found_value

    # --- Generic Handler for ANY Data Source Type ---
    
    # Path A: Column-Oriented (Dict of Lists) - e.g., processed_tables
    if isinstance(data_source, dict):
        # Determine number of rows from the first list found
        num_data_rows_from_source = 0
        for val in data_source.values():
            if isinstance(val, list):
                num_data_rows_from_source = max(num_data_rows_from_source, len(val))
        
        logger.debug(f"Preparing {num_data_rows_from_source} rows from Column-Oriented source")

        for i in range(num_data_rows_from_source):
            row_dict = {}
            for source_key, rule in dynamic_mapping_rules.items():
                if not isinstance(rule, dict): continue
                
                target_id = rule.get("column") or rule.get("id")
                if not target_id: continue
                target_col_idx = column_id_map.get(target_id)
                if not target_col_idx: continue
                
                # Fetch value using smart lookup (passing main dict and row index)
                val = get_value_from_row_or_cols(data_source, rule, source_key, row_idx=i)
                
                if val is not None:
                    row_dict[target_col_idx] = val
                
                # Apply Fallback
                current_val = row_dict.get(target_col_idx)
                if current_val in [None, ""]:
                    _apply_fallback(row_dict, target_col_idx, rule, DAF_mode)
            
            # Apply static values
            for col_idx, static_val in static_value_map.items():
                if col_idx not in row_dict:
                    row_dict[col_idx] = static_val
            
            data_rows_prepared.append(row_dict)

    # Path B: Row-Oriented (List of Dicts) - e.g., standard_aggregation_results
    elif isinstance(data_source, list):
        num_data_rows_from_source = len(data_source)
        logger.debug(f"Preparing {num_data_rows_from_source} rows from Row-Oriented source")
        
        for row_data in data_source:
            row_dict = {}
            for source_key, rule in dynamic_mapping_rules.items():
                if not isinstance(rule, dict): continue
                
                target_id = rule.get("column") or rule.get("id")
                if not target_id: continue
                target_col_idx = column_id_map.get(target_id)
                if not target_col_idx: continue
                
                # Fetch value using smart lookup (passing row object, no row index)
                val = get_value_from_row_or_cols(row_data, rule, source_key, row_idx=None)
                
                if val is not None:
                    row_dict[target_col_idx] = val

                # Apply Fallback
                current_val = row_dict.get(target_col_idx)
                if current_val in [None, ""]:
                    _apply_fallback(row_dict, target_col_idx, rule, DAF_mode)
            
            # Apply static values
            for col_idx, static_val in static_value_map.items():
                if col_idx not in row_dict:
                    row_dict[col_idx] = static_val
            
            data_rows_prepared.append(row_dict)
    
    else:
        logger.warning(f"Unknown data_source format: {type(data_source)}. Expected dict or list.")
    
    # Pad with empty rows if static labels demand it
    if num_static_labels > len(data_rows_prepared):
        data_rows_prepared.extend([{}] * (num_static_labels - len(data_rows_prepared)))
    
    return data_rows_prepared, pallet_counts_for_rows, dynamic_desc_used, num_data_rows_from_source