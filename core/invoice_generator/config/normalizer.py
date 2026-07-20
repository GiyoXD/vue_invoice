import copy
import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


def normalize_styling(sheet_styling: Dict[str, Any], defaults: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalize styling configuration for a sheet, converting legacy formats
    and applying global default border overrides.
    """
    if not isinstance(sheet_styling, dict):
        return {}

    # NEW FORMAT (columns + row_contexts)
    if 'columns' in sheet_styling and 'row_contexts' in sheet_styling:
        columns_copy = {col_id: col_def.copy() for col_id, col_def in sheet_styling['columns'].items()}
        
        borders = (defaults or {}).get('borders') or {}
        default_border = borders.get('default_border', 'full_grid')
        border_exceptions = borders.get('exceptions') or {}
        
        if isinstance(border_exceptions, dict):
            for col_id, border_style in border_exceptions.items():
                if col_id in columns_copy:
                    normalized_style = "sides_only" if border_style == "side_only" else border_style
                    columns_copy[col_id]['border_style'] = normalized_style
                    logger.debug(f"Merged border exception: {col_id} -> {normalized_style}")
        
        return {
            'columns': columns_copy,
            'row_contexts': sheet_styling['row_contexts'],
            'default_border': default_border
        }
    
    # OLD FORMAT transformation
    transformed = {}
    
    if 'header' in sheet_styling:
        header_cfg = sheet_styling['header']
        if 'font' in header_cfg:
            transformed['header_font'] = header_cfg['font']
        if 'alignment' in header_cfg:
            transformed['header_alignment'] = header_cfg['alignment']
        if 'row_height' in header_cfg:
            transformed.setdefault('row_heights', {})['header'] = header_cfg['row_height']
            
    if 'data' in sheet_styling:
        data_cfg = sheet_styling['data']
        if 'font' in data_cfg:
            transformed['default_font'] = data_cfg['font']
        if 'alignment' in data_cfg:
            transformed['default_alignment'] = data_cfg['alignment']
        if 'row_height' in data_cfg:
            transformed.setdefault('row_heights', {})['data_default'] = data_cfg['row_height']

    if 'footer' in sheet_styling:
        footer_cfg = sheet_styling['footer']
        if 'row_height' in footer_cfg:
            transformed.setdefault('row_heights', {})['footer'] = footer_cfg['row_height']

    if 'column_specific' in sheet_styling:
        transformed['column_id_styles'] = dict(sheet_styling['column_specific'])

    if 'dimensions' in sheet_styling:
        dims = sheet_styling['dimensions']
        if 'column_widths' in dims:
            transformed['column_id_widths'] = dims['column_widths']

    if 'borders' in (defaults or {}):
        transformed['borders'] = defaults['borders']

    return transformed


def normalize_layout(sheet_config: Dict[str, Any], defaults: Dict[str, Any]) -> Dict[str, Any]:
    """
    Deep-merge defaults into layout configuration for a sheet.
    """
    merged_sheet_config = copy.deepcopy(sheet_config or {})
    defaults = defaults or {}

    # Merge data_flow mappings
    default_mappings = defaults.get('data_flow', {}).get('mappings', {})
    if default_mappings:
        sheet_data_flow = merged_sheet_config.get('data_flow', {})
        sheet_mappings = sheet_data_flow.get('mappings', {})

        merged_mappings = {}
        all_keys = set(default_mappings.keys()) | set(sheet_mappings.keys())
        for key in all_keys:
            default_rule = default_mappings.get(key, {})
            sheet_rule = sheet_mappings.get(key, {})
            if isinstance(default_rule, dict) and isinstance(sheet_rule, dict):
                merged_mappings[key] = {**default_rule, **sheet_rule}
            elif key in sheet_mappings:
                merged_mappings[key] = sheet_rule
            else:
                merged_mappings[key] = default_rule

        merged_sheet_config.setdefault('data_flow', {})['mappings'] = merged_mappings

    # Merge footer
    default_footer = defaults.get('footer', {})
    if default_footer:
        sheet_footer = merged_sheet_config.get('footer', {})
        for key, default_val in default_footer.items():
            if key not in sheet_footer:
                sheet_footer[key] = copy.deepcopy(default_val)
        merged_sheet_config['footer'] = sheet_footer

    return merged_sheet_config
