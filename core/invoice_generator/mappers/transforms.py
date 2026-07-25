import ast
import re
import logging
from typing import Any, Dict, List, Optional, Tuple, Union
from .rules import apply_fallback, parse_formula_def, resolve_mode_formula

logger = logging.getLogger(__name__)


def extract_table_data(data_source: Any, data_source_type: str) -> Any:
    """Extract table data for processing, resolving stringified tuple keys if present."""
    if data_source is None:
        return None
    
    if data_source_type in ['processed_tables', 'processed_tables_multi']:
        return data_source
    
    if isinstance(data_source, dict):
        new_data = {}
        for k, v in data_source.items():
            if isinstance(k, str) and k.startswith('(') and k.endswith(')'):
                try:
                    clean_k = re.sub(r"Decimal\((['\"])(.*?)\1\)", r"\2", k)
                    new_key = ast.literal_eval(clean_k)
                    new_data[new_key] = v
                except (ValueError, SyntaxError):
                    new_data[k] = v
            else:
                new_data[k] = v
        return new_data

    return data_source


def _get_value_from_source(
    source_container: Any,
    rule_key: str,
    rule: Optional[Dict[str, Any]] = None,
    row_idx: Optional[int] = None
) -> Any:
    """Extract a value from data source using explicit column/field mapping or rule_key."""
    if not isinstance(source_container, dict):
        return None

    lookup_keys = []
    if isinstance(rule, dict):
        for k in ['column', 'field', 'source', 'key']:
            mapped_key = rule.get(k)
            if mapped_key and mapped_key not in lookup_keys:
                lookup_keys.append(mapped_key)
    if rule_key and rule_key not in lookup_keys:
        lookup_keys.append(rule_key)

    for key in lookup_keys:
        if row_idx is not None:
            col_data = source_container.get(key)
            if isinstance(col_data, list) and row_idx < len(col_data):
                val = col_data[row_idx]
                if val not in [None, ""]:
                    return val
            elif col_data not in [None, ""] and not isinstance(col_data, list):
                return col_data
        elif key in source_container and source_container[key] not in [None, ""]:
            return source_container[key]

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
    """Build a single row dictionary by applying mapping rules, formulas, fallbacks, and static values."""
    row_dict = {}

    for source_key, rule in dynamic_mapping_rules.items():
        if not isinstance(rule, dict):
            continue

        target_id = source_key
        if not target_id:
            continue
        if target_id in parent_column_ids:
            continue

        val = _get_value_from_source(source_container, source_key, rule, row_idx)
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

    for col_id, static_val in static_value_map.items():
        row_dict[col_id] = static_val

    return row_dict


def prepare_data_rows(
    data_source_type: str,
    data_source: Any,
    dynamic_mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str],
    desc_col_idx: int,
    num_static_labels: int,
    static_value_map: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool,
    parent_column_ids: List[str],
    pricing_net_weight: bool = False
) -> Tuple[List[Dict[str, Any]], int]:
    """Prepare formatted data rows for table rendering."""
    data_rows = []

    if isinstance(data_source, dict):
        num_rows = 0
        for val in data_source.values():
            if isinstance(val, list):
                num_rows = max(num_rows, len(val))
                    
        for r_idx in range(num_rows):
            row_dict = _build_row_dict(
                source_container=data_source,
                row_idx=r_idx,
                dynamic_mapping_rules=dynamic_mapping_rules,
                parent_column_ids=parent_column_ids,
                static_value_map=static_value_map,
                DAF_mode=DAF_mode,
                custom_mode=custom_mode,
                pricing_net_weight=pricing_net_weight
            )
            data_rows.append(row_dict)
    elif isinstance(data_source, list):
        for item in data_source:
            if isinstance(item, dict):
                row_dict = _build_row_dict(
                    source_container=item,
                    row_idx=None,
                    dynamic_mapping_rules=dynamic_mapping_rules,
                    parent_column_ids=parent_column_ids,
                    static_value_map=static_value_map,
                    DAF_mode=DAF_mode,
                    custom_mode=custom_mode,
                    pricing_net_weight=pricing_net_weight
                )
                data_rows.append(row_dict)

    num_data_rows = len(data_rows)
    return data_rows, num_data_rows


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
    if not static_payload or 'col_static' not in static_payload:
        return

    static_values = static_payload['col_static']
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
    if num_data_rows <= 0 or not pallet_col_id:
        return

    carry_value = None
    for row in data_rows[:num_data_rows]:
        raw_val = row.get(pallet_col_id)
        if raw_val not in (None, "", 0):
            carry_value = raw_val
        elif carry_value is not None:
            row[pallet_col_id] = carry_value


def extract_summaries(
    data_source: Any,
    footer_data: Dict[str, Any],
    table_key: Optional[str] = None
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], Optional[int]]:
    """Extract leather_summary, weight_summary, and pallet_summary_total."""
    leather_summary = None
    weight_summary = None
    pallet_summary_total = None
    
    if isinstance(data_source, dict):
        leather_summary = data_source.get('leather_summary')
        if not leather_summary and footer_data and 'add_ons' in footer_data:
            leather_summary = footer_data['add_ons'].get('leather_summary_addon')
        
        weight_summary = data_source.get('weight_summary')
        if not weight_summary and footer_data and 'add_ons' in footer_data:
            weight_summary = footer_data['add_ons'].get('weight_summary_addon')
    elif isinstance(data_source, list):
        if footer_data and 'add_ons' in footer_data:
            leather_summary = footer_data['add_ons'].get('leather_summary_addon')
            weight_summary = footer_data['add_ons'].get('weight_summary_addon')
            
    if footer_data:
        if 'grand_total' in footer_data and 'col_pallet_count' in footer_data['grand_total'] and table_key is None:
            pallet_summary_total = int(footer_data['grand_total']['col_pallet_count'])
        
        if pallet_summary_total is None and 'table_totals' in footer_data:
            table_totals = footer_data['table_totals']
            idx = 0
            if table_key is not None:
                try:
                    idx = int(table_key)
                except ValueError:
                    idx = 0
            if isinstance(table_totals, list) and 0 <= idx < len(table_totals):
                tbl_footer = table_totals[idx]
                if 'col_pallet_count' in tbl_footer:
                    pallet_summary_total = int(tbl_footer['col_pallet_count'])
            elif isinstance(table_totals, dict):
                key = str(table_key) if table_key is not None else next(iter(table_totals.keys()), "")
                first_val = table_totals.get(key, next(iter(table_totals.values()), {}))
                if 'col_pallet_count' in first_val:
                    pallet_summary_total = int(first_val['col_pallet_count'])
                    
    if pallet_summary_total is None and isinstance(data_source, dict):
        pallet_summary_total = data_source.get('pallet_summary_total')
        if pallet_summary_total is not None:
            logger.warning(f"Using legacy pallet_summary_total from data_source: {pallet_summary_total}")
            
    return leather_summary, weight_summary, pallet_summary_total


__all__ = [
    "extract_table_data",
    "prepare_data_rows",
    "resolve_static_placeholders",
    "populate_static_content",
    "format_pallet_counts",
    "extract_summaries"
]
