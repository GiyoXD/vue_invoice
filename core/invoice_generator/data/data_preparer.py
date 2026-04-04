from typing import Any, Union, Dict, List, Tuple, Optional
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

        # For all other rules, strictly use the rule_key as the target ID
        target_id = rule_key
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
        # Preserve precision: whole numbers → int, fractional → float via string
        # to avoid IEEE 754 artifacts (e.g. 0.30000000000000004)
        if value == value.to_integral_value():
            return int(value)
        return float(str(value))
    return value # Return original value for other types

def _apply_fallback(
    row_dict: Dict[int, Any],
    target_col_idx: int,
    mapping_rule: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool
):
    """
    Applies a fallback value to the row_dict based on the DAF_mode and custom_mode.
    
    Supports:
    1. Modern nested format: "fallback": {"standard": "X", "daf": "Y", "custom": "Z", "default": "W"}
       Resolution order: mode-specific (daf/custom) → standard → default
    2. Legacy flat format: "fallback": "X" (same value for all modes)
    """
    # Priority 1: Check Modern Nested Dictionary Structure
    fallback_config = mapping_rule.get('fallback')
    if isinstance(fallback_config, dict):
        if DAF_mode and 'daf' in fallback_config:
            row_dict[target_col_idx] = fallback_config['daf']
            return
        elif custom_mode and 'custom' in fallback_config:
            row_dict[target_col_idx] = fallback_config['custom']
            return
        elif 'standard' in fallback_config:
            row_dict[target_col_idx] = fallback_config['standard']
            return
        elif 'default' in fallback_config:
            # Universal catch-all: applies to any mode when no specific key matches
            row_dict[target_col_idx] = fallback_config['default']
            return
    elif fallback_config is not None:
        # Priority 2: Try single 'fallback' string key (same value for all modes)
        row_dict[target_col_idx] = fallback_config
        return

import re

def _parse_formula_def(formula_def: Union[str, Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """
    Normalizes a formula definition into a dictionary with 'template' and 'inputs'.
    Supports both legacy dict format: {"template": "{col_a} * {col_b}", "inputs": ["col_a", "col_b"]}
    and modern string format: "{col_a} * {col_b}" (auto-extracts inputs via regex).
    """
    if isinstance(formula_def, dict) and 'template' in formula_def:
        return formula_def
        
    if isinstance(formula_def, str) and formula_def.strip():
        # Auto-extract inputs encased in curly brackets {like_this}
        inputs = re.findall(r'\{([^}]+)\}', formula_def)
        # Filter out special non-column placeholders like {row}
        filtered_inputs = [inp for inp in inputs if inp != 'row']
        
        # Rewrite the template to use {col_ref_0}, {col_ref_1}, etc. format
        # This is strictly required by the data_table_builder.py engine
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

    Supports:
    1. Modern nested dict: "formula": {"standard": "...", "daf": "...", "custom": "..."}
    2. Plain string: "formula": "{col_a} * {col_b}" (applies to all modes)

    Args:
        rule: The mapping rule dict for a column.
        DAF_mode: Whether the current invoice is in DAF mode.
        custom_mode: Whether the current invoice is in custom mode.

    Returns:
        The formula string for the current mode, or None if no formula applies.
    """
    formula_config = rule.get('formula')
    if formula_config is None:
        return None

    # Modern nested dict: pick by mode
    if isinstance(formula_config, dict):
        if custom_mode and 'custom' in formula_config:
            return formula_config['custom']
        elif DAF_mode and 'daf' in formula_config:
            return formula_config['daf']
        elif not custom_mode and not DAF_mode and 'standard' in formula_config:
            return formula_config['standard']
        return None  # Dict exists but no key matches this mode

    # Plain string: applies to all modes
    if isinstance(formula_config, str):
        return formula_config

    return None

def _get_value_from_source(source_container: Any, rule_key: str, row_idx: int = None) -> Any:
    """
    Extract a value from the data source using the rule key.

    Args:
        source_container: Either the main data dict (column-oriented) or a single row dict (row-oriented).
        rule_key: The column ID to look up (e.g., 'col_po', 'col_desc').
        row_idx: Row index for column-oriented sources. None for row-oriented.

    Returns:
        The extracted value, or None if not found.
    """
    if rule_key is None or not isinstance(source_container, dict):
        return None

    # Column-Oriented: source is dict of lists, need row_idx
    if row_idx is not None:
        col_data = source_container.get(rule_key)
        if isinstance(col_data, list) and row_idx < len(col_data):
            return col_data[row_idx]

    # Row-Oriented: source is a single row dict
    elif rule_key in source_container:
        return source_container[rule_key]

    return None


def _build_row_dict(
    source_container: Any,
    row_idx: Optional[int],
    dynamic_mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    parent_column_ids: List[str],
    static_value_map: Dict[int, Any],
    DAF_mode: bool,
    custom_mode: bool
) -> Dict[int, Any]:
    """
    Build a single row dictionary by applying mapping rules, formulas, fallbacks, and static values.

    Args:
        source_container: The data source (full dict for column-oriented, single row dict for row-oriented).
        row_idx: Row index for column-oriented lookups. None for row-oriented.
        dynamic_mapping_rules: Parsed mapping rules from the config.
        column_id_map: Column ID to Excel column index mapping.
        parent_column_ids: IDs of parent (merged header) columns to skip.
        static_value_map: Static values to inject into each row.
        DAF_mode: Whether DAF mode is active.
        custom_mode: Whether custom mode is active.

    Returns:
        A dictionary mapping column indices to their resolved values.
    """
    row_dict = {}

    for source_key, rule in dynamic_mapping_rules.items():
        if not isinstance(rule, dict):
            continue

        target_id = source_key
        if not target_id:
            continue
        # Skip parent columns since data should only be written to leaf columns
        if target_id in parent_column_ids:
            continue

        target_col_idx = column_id_map.get(target_id)
        if not target_col_idx:
            continue

        # Fetch value from source
        val = _get_value_from_source(source_container, source_key, row_idx)
        if val is not None:
            row_dict[target_col_idx] = val

        # Formula-First Resolution: formula always wins over raw data
        mode_formula = _resolve_mode_formula(rule, DAF_mode, custom_mode)
        if mode_formula:
            parsed_formula = _parse_formula_def(mode_formula)
            if parsed_formula:
                row_dict[target_col_idx] = {
                    'type': 'formula',
                    'template': parsed_formula['template'],
                    'inputs': parsed_formula.get('inputs', [])
                }
                continue  # Formula applied, skip fallback

        # No formula for this mode — apply text fallback if value is empty
        if row_dict.get(target_col_idx) in [None, ""]:
            # Guard: Only apply col_desc fallback to rows that have actual data (e.g., a PO value)
            # IMPORTANT: Check SOURCE DATA directly, not row_dict, because col_po
            # may not have been processed yet in this iteration loop.
            if source_key == 'col_desc':
                po_val = _get_value_from_source(source_container, 'col_po', row_idx)
                if not po_val:
                    continue  # Skip fallback — this is an overflow row with no real data
            _apply_fallback(row_dict, target_col_idx, rule, DAF_mode, custom_mode)

    # Apply static values
    for col_idx, static_val in static_value_map.items():
        if col_idx not in row_dict:
            row_dict[col_idx] = static_val

    return row_dict


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
    custom_mode: bool = False,
    parent_column_ids: List[str] = None
) -> Tuple[List[Dict[int, Any]], List[int], int]:
    """
    Prepares data rows by applying mapping rules to the data source.
    Supports both Column-Oriented (Dict of Lists) and Row-Oriented (List of Dicts) sources.
    """
    parent_column_ids = parent_column_ids or []
    data_rows_prepared = []
    pallet_counts_for_rows = []
    num_data_rows_from_source = 0

    # Shared kwargs for _build_row_dict
    build_kwargs = {
        'dynamic_mapping_rules': dynamic_mapping_rules,
        'column_id_map': column_id_map,
        'parent_column_ids': parent_column_ids,
        'static_value_map': static_value_map,
        'DAF_mode': DAF_mode,
        'custom_mode': custom_mode,
    }

    # Path A: Column-Oriented (Dict of Lists) - e.g., processed_tables
    if isinstance(data_source, dict):
        # Determine number of rows from the longest list
        for val in data_source.values():
            if isinstance(val, list):
                num_data_rows_from_source = max(num_data_rows_from_source, len(val))

        logger.debug(f"Preparing {num_data_rows_from_source} rows from Column-Oriented source")

        # Extract pallet counts
        raw_pallet_counts = data_source.get('col_pallet_count')
        if raw_pallet_counts and isinstance(raw_pallet_counts, list):
            pallet_counts_for_rows = raw_pallet_counts
        else:
            logger.warning(f"[DataPreparer] ⚠️ Pallet count missing in column-oriented source. Defaulting to None for {num_data_rows_from_source} rows.")
            pallet_counts_for_rows = [None] * num_data_rows_from_source

        for i in range(num_data_rows_from_source):
            row_dict = _build_row_dict(source_container=data_source, row_idx=i, **build_kwargs)
            data_rows_prepared.append(row_dict)

    # Path B: Row-Oriented (List of Dicts) - e.g., standard_aggregation_results
    elif isinstance(data_source, list):
        num_data_rows_from_source = len(data_source)
        logger.debug(f"Preparing {num_data_rows_from_source} rows from Row-Oriented source")

        for row_data in data_source:
            # Extract pallet count per row
            p_count = row_data.get('col_pallet_count') or row_data.get('pallet_count')
            pallet_counts_for_rows.append(p_count)

            row_dict = _build_row_dict(source_container=row_data, row_idx=None, **build_kwargs)
            logger.debug(f"[DEBUG-ROW] Source keys: {list(row_data.keys())} - target row dict for row {len(data_rows_prepared)}: {row_dict}")
            data_rows_prepared.append(row_dict)

    else:
        logger.warning(f"Unknown data_source format: {type(data_source)}. Expected dict or list.")

    # Pad with empty rows if static labels demand it
    if num_static_labels > len(data_rows_prepared):
        data_rows_prepared.extend([{}] * (num_static_labels - len(data_rows_prepared)))

    return data_rows_prepared, pallet_counts_for_rows, num_data_rows_from_source