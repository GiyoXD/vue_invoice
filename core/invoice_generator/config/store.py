import copy
import logging
from functools import lru_cache
from typing import Any, Dict, List, Optional

from .normalizer import normalize_layout, normalize_styling

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
        self.raw_config = config_data or {}
        
        # Extract metadata
        meta = self.raw_config.get('_meta', {})
        self.version = meta.get('config_version', 'unknown')
        self.customer = meta.get('customer', 'unknown')
        
        # Parsed sections
        self._processing = self.raw_config.get('processing', {})
        self._styling_bundle = self.raw_config.get('styling_bundle', {})
        self._layout_bundle = self.raw_config.get('layout_bundle', {})
        self._data_bundle = self.raw_config.get('data_bundle', {})
        
        if template_data and "template_layout" in template_data:
            self.template_json_config = template_data["template_layout"]
        else:
            self.template_json_config = template_data
            
        logger.info(f"ConfigStore initialized successfully. Version: {self.version}")
        
    def get_sheets_to_process(self) -> List[str]:
        """Get list of sheets to process."""
        return list(self._processing.keys())
    
    def get_data_source_type(self, sheet_name: str, mode: str = "standard") -> Optional[str]:
        """Get data source type for a sheet, resolved by mode if config is mode-aware."""
        val = self._processing.get(sheet_name)
        if isinstance(val, dict):
            return val.get(mode, val.get("standard"))
        return val
    
    def get_sheet_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get complete config for a sheet (combines all bundles).
        """
        return {
            'data_source': self._processing.get(sheet_name),
            'styling_config': self.get_styling_config(sheet_name),
            'layout_config': self.get_layout_config(sheet_name),
            'data_config': self.get_data_config(sheet_name)
        }
    
    def get_styling_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get styling configuration for a sheet, transformed and normalized.
        """
        sheet_styling = self._styling_bundle.get(sheet_name, {})
        defaults = self._styling_bundle.get('defaults', {})
        return normalize_styling(sheet_styling, defaults)
    
    @staticmethod
    @lru_cache(maxsize=1)
    def _get_master_defaults() -> Dict[str, Any]:
        """
        Load defaults section from master_config.json as global fallback.
        
        NOTE: This is cached in process memory via @lru_cache.
        If master_config.json on disk is modified while the dev server is running,
        you MUST restart the dev server (.\start_dev.ps1) or call
        ConfigStore._get_master_defaults.cache_clear() for changes to take effect.
        """
        try:
            from core.system_config import sys_config
            import json
            master_path = sys_config.blueprints_root / "mapper" / "master_config.json"
            if master_path.exists():
                with open(master_path, 'r', encoding='utf-8') as f:
                    master_data = json.load(f)
                    return master_data.get('layout_bundle', {}).get('defaults', {})
        except Exception as e:
            logger.warning(f"Could not load master_config.json defaults: {e}")
        return {}

    def get_layout_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get layout configuration for a sheet (headers, blanks, static content, merges),
        with global defaults merged in.
        """
        sheet_config = self._layout_bundle.get(sheet_name, {})
        defaults = copy.deepcopy(self._layout_bundle.get('defaults', {}))
        master_defaults = self._get_master_defaults()
        
        # Merge master defaults for any missing default keys (like static_payload, mappings, footer)
        if master_defaults:
            for k, v in master_defaults.items():
                if k not in defaults:
                    defaults[k] = copy.deepcopy(v)

        return normalize_layout(sheet_config, defaults)
    
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
        """Get the raw config dictionary."""
        return self.raw_config
