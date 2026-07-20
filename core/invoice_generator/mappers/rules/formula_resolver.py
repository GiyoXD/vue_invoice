import re
from typing import Any, Dict, Optional, Union


def parse_formula_def(formula_def: Union[str, Dict[str, Any]]) -> Optional[Dict[str, Any]]:
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


def resolve_mode_formula(
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
