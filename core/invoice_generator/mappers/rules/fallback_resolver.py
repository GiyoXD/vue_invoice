from typing import Any, Dict


def apply_fallback(
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
