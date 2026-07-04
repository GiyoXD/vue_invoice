import copy
import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class ConfigStore:
    """
    Manages the state and provides query interfaces for loaded bundled configuration and template data.
    Does not perform any File I/O.
    """
    
    def __init__(self, config_data: Dict[str, Any], template_data: Optional[Dict[str, Any]] = None):
        """
        Initialize the ConfigStore with configuration and template dictionaries.
        
        Args:
            config_data: Parsed configuration dictionary.
            template_data: Parsed template layout dictionary (optional).
        """
        self.raw_config = config_data
        
        # Extract metadata
        meta = config_data.get('_meta', {})
        self.version = meta.get('config_version', 'unknown')
        self.customer = meta.get('customer', 'unknown')
        
        # Parsed sections
        self._processing = config_data.get('processing', {})
        self._styling_bundle = config_data.get('styling_bundle', {})
        self._layout_bundle = config_data.get('layout_bundle', {})
        self._data_bundle = config_data.get('data_bundle', {})
        if template_data and "template_layout" in template_data:
            self.template_json_config = template_data["template_layout"]
        else:
            self.template_json_config = template_data
        
        logger.info(f"ConfigStore initialized successfully. Version: {self.version}")
        
    def get_sheets_to_process(self) -> List[str]:
        """Get list of sheets to process."""
        return list(self._processing.keys())
    
    def get_data_source_type(self, sheet_name: str) -> Optional[str]:
        """
        Get data source type for a sheet.
        
        Returns:
            'aggregation', 'DAF_aggregation', 'processed_tables_multi', etc.
        """
        return self._processing.get(sheet_name)
    
    def get_sheet_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get complete config for a sheet (combines all bundles).
        
        This is the main method processors should use to get sheet configuration.
        Returns a unified config dictionary with all the needed sections.
        """
        return {
            'data_source': self.get_data_source_type(sheet_name),
            'styling_config': self.get_styling_config(sheet_name),
            'layout_config': self.get_layout_config(sheet_name),
            'data_config': self.get_data_config(sheet_name)
        }
    
    def get_styling_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get styling configuration for a sheet, transformed to StylingConfigModel format.
        
        Transforms bundled config format:
            {"header": {"font": {...}}, "data": {"font": {...}}}
        Into StylingConfigModel format:
            {"header_font": {...}, "default_font": {...}}
        
        OR if new format is detected (columns + row_contexts), returns them as-is.
        """
        # Get sheet-specific styling
        sheet_styling = self._styling_bundle.get(sheet_name, {})
        
        # DEBUG: Log what we're checking
        logger.debug(f"get_styling_config for '{sheet_name}'")
        logger.debug(f"Keys in sheet_styling: {list(sheet_styling.keys()) if isinstance(sheet_styling, dict) else 'NOT A DICT'}")
        
        # Check if using NEW FORMAT (columns + row_contexts)
        if 'columns' in sheet_styling and 'row_contexts' in sheet_styling:
            # New format: return as-is, don't transform
            logger.debug(f"NEW FORMAT detected - returning columns + row_contexts as-is")
            
            # Create a copy of the columns dictionary to avoid mutating cached configuration in-place
            columns_copy = {col_id: col_def.copy() for col_id, col_def in sheet_styling['columns'].items()}
            
            # Extract and merge border exceptions from global defaults
            defaults = self._styling_bundle.get('defaults') or {}
            borders = defaults.get('borders') or {}
            default_border = borders.get('default_border', 'full_grid')
            border_exceptions = borders.get('exceptions') or {}
            if isinstance(border_exceptions, dict):
                for col_id, border_style in border_exceptions.items():
                    if col_id in columns_copy:
                        # Normalize "side_only" from config to "sides_only" expected by cells
                        normalized_style = "sides_only" if border_style == "side_only" else border_style
                        columns_copy[col_id]['border_style'] = normalized_style
                        logger.debug(f"Merged border exception: {col_id} -> {normalized_style}")
            
            return {
                'columns': columns_copy,
                'row_contexts': sheet_styling['row_contexts'],
                'default_border': default_border
            }
        
        # OLD FORMAT: Transform nested bundled format to flat StylingConfigModel format
        logger.debug(f"OLD FORMAT detected - transforming to StylingConfigModel format")
        # Get default styling to use as fallback
        defaults = self._styling_bundle.get('defaults', {})
        
        # Transform nested bundled format to flat StylingConfigModel format
        transformed = {}
        
        # Extract header styling
        if 'header' in sheet_styling:
            header_cfg = sheet_styling['header']
            if 'font' in header_cfg:
                transformed['header_font'] = header_cfg['font']
            if 'alignment' in header_cfg:
                transformed['header_alignment'] = header_cfg['alignment']
            if 'row_height' in header_cfg:
                if 'row_heights' not in transformed:
                    transformed['row_heights'] = {}
                transformed['row_heights']['header'] = header_cfg['row_height']
        
        # Extract data (default) styling
        if 'data' in sheet_styling:
            data_cfg = sheet_styling['data']
            if 'font' in data_cfg:
                transformed['default_font'] = data_cfg['font']
            if 'alignment' in data_cfg:
                transformed['default_alignment'] = data_cfg['alignment']
            if 'row_height' in data_cfg:
                if 'row_heights' not in transformed:
                    transformed['row_heights'] = {}
                transformed['row_heights']['data_default'] = data_cfg['row_height']
        
        # Extract footer styling
        if 'footer' in sheet_styling:
            footer_cfg = sheet_styling['footer']
            if 'row_height' in footer_cfg:
                if 'row_heights' not in transformed:
                    transformed['row_heights'] = {}
                transformed['row_heights']['footer'] = footer_cfg['row_height']
        
        # Extract column-specific styling
        if 'column_specific' in sheet_styling:
            col_styles = {}
            for col_id, col_cfg in sheet_styling['column_specific'].items():
                col_styles[col_id] = col_cfg
            transformed['column_id_styles'] = col_styles
        
        # Extract dimensions (column widths)
        if 'dimensions' in sheet_styling:
            dims = sheet_styling['dimensions']
            if 'column_widths' in dims:
                transformed['column_id_widths'] = dims['column_widths']
        
        # Extract border configuration from defaults
        if 'borders' in defaults:
            # Borders are handled separately, just pass through
            transformed['borders'] = defaults['borders']
        
        return transformed
    
    def get_layout_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get layout configuration for a sheet (headers, blanks, static content, merges).

        Merges layout_bundle.defaults as a base layer for both data_flow.mappings
        and footer config. Per-sheet values override defaults.

        Args:
            sheet_name: Name of the sheet to get config for.

        Returns:
            The sheet's layout config with defaults merged in.
        """
        sheet_config = copy.deepcopy(self._layout_bundle.get(sheet_name, {}))

        defaults = self._layout_bundle.get('defaults', {})

        # --- Merge defaults.data_flow.mappings (DEEP MERGE) ---
        default_mappings = defaults.get('data_flow', {}).get('mappings', {})
        if default_mappings:
            sheet_data_flow = sheet_config.get('data_flow', {})
            sheet_mappings = sheet_data_flow.get('mappings', {})

            # Deep merge: for each column, merge default rule with per-sheet rule.
            # Per-sheet keys take priority, but an empty {} doesn't wipe out defaults.
            merged = {}
            all_keys = set(default_mappings.keys()) | set(sheet_mappings.keys())
            for key in all_keys:
                default_rule = default_mappings.get(key, {})
                sheet_rule = sheet_mappings.get(key, {})
                if isinstance(default_rule, dict) and isinstance(sheet_rule, dict):
                    # Merge: default as base, sheet overrides on top
                    merged[key] = {**default_rule, **sheet_rule}
                elif key in sheet_mappings:
                    merged[key] = sheet_rule  # Per-sheet non-dict overrides entirely
                else:
                    merged[key] = default_rule  # Only in defaults

            sheet_config.setdefault('data_flow', {})['mappings'] = merged

            logger.debug(
                f"[{sheet_name}] Deep-merged {len(default_mappings)} default mappings + "
                f"{len(sheet_mappings)} sheet mappings = {len(merged)} total"
            )

        # --- Merge defaults.footer ---
        default_footer = defaults.get('footer', {})
        if default_footer:
            sheet_footer = sheet_config.get('footer', {})

            # For each default footer key, use it as base if not in per-sheet
            for key, default_val in default_footer.items():
                if key not in sheet_footer:
                    sheet_footer[key] = copy.deepcopy(default_val)
                    logger.debug(
                        f"[{sheet_name}] Inherited footer.{key} from defaults"
                    )

            sheet_config['footer'] = sheet_footer

        return sheet_config
    
    def get_data_config(self, sheet_name: str) -> Dict[str, Any]:
        """Get data configuration for a sheet (mappings, header_info, etc.)."""
        return self._data_bundle.get(sheet_name, {})
        
    def get_template_json_config(self) -> Optional[Dict[str, Any]]:
        """Get the loaded sibling template JSON config if available."""
        return self.template_json_config
    
    def has_static_sheets(self) -> bool:
        """Check if the config explicitly indicates presence of static sheets."""
        meta = self.raw_config.get('_meta', {})
        return meta.get('has_static_sheets', False)

    def is_bundled_config(self) -> bool:
        """Check if this is a bundled config (v2.1+)."""
        return self.version.startswith('2.1')
    
    def get_max_columns(self, sheet_name: str) -> Optional[int]:
        """
        Count the actual number of Excel columns from the template header or layout config.
        Accounts for parent columns with children if falling back to layout config.
        """
        template_config = self.get_template_json_config()
        if template_config and sheet_name in template_config:
            sheet_layout = template_config[sheet_name]
            from core.blueprint_generator.internal.scanner.models import TemplateLayout
            layout = TemplateLayout.from_dict(sheet_layout)
            if layout:
                max_col = 0
                for row in layout.header_rows:
                    for cell in row.cells:
                        max_col = max(max_col, cell.col_index)
                        if cell.merge:
                            max_col = max(max_col, cell.merge.max_col)
                if max_col > 0:
                    logger.debug(f"[PrintArea] Safe max column {max_col} from template JSON for '{sheet_name}'")
                    return max_col

        layout = self.get_layout_config(sheet_name)
        columns = layout.get('structure', {}).get('columns', [])
        if not columns:
            return None

        count = 0
        for col in columns:
            children = col.get('children', [])
            if children:
                count += len(children)
            else:
                count += 1

        logger.debug(f"[PrintArea] Layout column count for '{sheet_name}': {count}")
        return count
    
    def get_raw_config(self) -> Dict[str, Any]:
        """Get the raw config dictionary (avoid using this if possible)."""
        return self.raw_config
