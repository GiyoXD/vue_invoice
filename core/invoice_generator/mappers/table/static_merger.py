import logging
from typing import Any, Dict, List

logger = logging.getLogger(__name__)


def merge_static_content(
    data_rows: List[Dict[str, Any]],
    static_content: Dict[str, Any],
    dynamic_mapping_rules: Dict[str, Any],
    DAF_mode: bool,
    custom_mode: bool
) -> None:
    """
    Merge static content into the first N data rows.
    Extends data_rows if there are more static values than data rows.
    """
    if not static_content or 'col_static' not in static_content:
        return

    static_values = static_content['col_static']
    static_col_id = 'col_static'
    
    if static_values and len(data_rows) > 0:
        # Resolve {col_desc_fallback} placeholder for dynamic static content
        desc_fallback_str = ""
        for rule_key, rule in dynamic_mapping_rules.items():
            if rule_key == 'col_desc' and isinstance(rule, dict):
                fallback_cfg = rule.get('fallback')
                if isinstance(fallback_cfg, dict):
                    if DAF_mode and 'daf' in fallback_cfg:
                        desc_fallback_str = fallback_cfg['daf']
                    elif custom_mode and 'custom' in fallback_cfg:
                        desc_fallback_str = fallback_cfg['custom']
                    elif 'standard' in fallback_cfg:
                        desc_fallback_str = fallback_cfg['standard']
                elif fallback_cfg is not None:
                    desc_fallback_str = str(fallback_cfg)
                break
        
        num_static_values = len(static_values)
        
        # Extend data_rows if we have more static values than data rows
        while len(data_rows) < num_static_values:
            data_rows.append({})

        for i, static_value in enumerate(static_values):
            if isinstance(static_value, str) and "{col_desc_fallback}" in static_value:
                static_value = static_value.replace("{col_desc_fallback}", str(desc_fallback_str))
                
            data_rows[i][static_col_id] = static_value
        
        logger.info(f"Merged {num_static_values} static values into {len(data_rows)} data rows")
