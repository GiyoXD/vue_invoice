# invoice_generator/config/config_loader.py
"""
Config Loader for Bundled Config Format (v2.1)

This module provides a clean interface to load and parse bundled configuration files.
It understands the v2.1 bundled format structure and provides easy access to sheet configs.

Bundled Config Structure:
    - _meta: metadata (version, customer, etc.)
    - processing: sheets list, data_sources
    - styling_bundle: per-sheet styling configs
    - layout_bundle: per-sheet layout configs (headers, blanks, static content)
    - data_bundle: per-sheet data configs (mappings, header_info)
    - context: global context (replacements, features, extensions)
"""

import json
from pathlib import Path
from typing import Any, Dict, Optional, List
import logging
logger = logging.getLogger(__name__)


class BundledConfigLoader:
    """
    Loads and parses bundled config files (v2.1+).
    
    Provides clean access to per-sheet configurations without polluting the main script.
    """
    
    def __init__(self, config_path: Path):
        """
        Initialize the config loader.
        
        Args:
            config_path: Path to the bundled config JSON file
        """
        self.config_path = config_path
        self.raw_config: Dict[str, Any] = {}
        self.version: str = "unknown"
        self.customer: str = "unknown"
        
        # Parsed sections
        self._processing: Dict[str, Any] = {}
        self._styling_bundle: Dict[str, Any] = {}
        self._layout_bundle: Dict[str, Any] = {}
        self._data_bundle: Dict[str, Any] = {}
        self._context: Dict[str, Any] = {}
        
        self._load()
    
    def _load(self) -> None:
        """Load and parse the config file."""
        logger.debug(f"Loading configuration from: {self.config_path}")
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.raw_config = json.load(f)
            
            # Extract metadata
            meta = self.raw_config.get('_meta', {})
            self.version = meta.get('config_version', 'unknown')
            self.customer = meta.get('customer', 'unknown')
            logger.info(f"Configuration loaded successfully.")
            logger.info(f"Detected bundled config version: {self.version}")
            
            # Parse main sections
            self._processing = self.raw_config.get('processing', {})
            self._styling_bundle = self.raw_config.get('styling_bundle', {})
            self._layout_bundle = self.raw_config.get('layout_bundle', {})
            self._data_bundle = self.raw_config.get('data_bundle', {})
            self._context = self.raw_config.get('context', {})
            
            # Load sibling template config for JSON-based reconstruction
            self.template_json_config: Dict[str, Any] = None
            try:
                # Deduce template json path: same dir, "{config_name}_template.json"
                # Convention: {CLIENT}_config.json -> {CLIENT}_template.json
                # OR just side by side replacement: _config.json -> _template.json
                stem = self.config_path.stem
                parent = self.config_path.parent
                
                # Try simple replacement if suffix exists
                if stem.endswith('_config'):
                    template_name = stem.replace('_config', '_template') + ".json"
                else:
                    template_name = f"{stem}_template.json"
                    
                template_path = parent / template_name
                if template_path.exists():
                    with open(template_path, 'r', encoding='utf-8') as f:
                        raw_tmpl = json.load(f)
                        # The file usually has root {"template_layout": {...}}
                        self.template_json_config = raw_tmpl.get("template_layout", {})
                        logger.info(f"Loaded sibling template config from: {template_path}")
                else:
                    logger.debug(f"No sibling template JSON found at {template_path}")
            except Exception as e:
                logger.warning(f"Failed to load sibling template JSON: {e}")

        except Exception as e:
            logger.error(f"Error loading configuration file {self.config_path}: {e}")
            raise
    
    # --- Public Interface ---
    
    def get_sheets_to_process(self) -> List[str]:
        """Get list of sheets to process."""
        # Support alias 'processing_order' for clarity
        return self._processing.get('sheets', self._processing.get('processing_order', []))
    
    def get_data_source_type(self, sheet_name: str) -> Optional[str]:
        """
        Get data source type for a sheet.
        
        Returns:
            'aggregation', 'DAF_aggregation', 'processed_tables_multi', etc.
        """
        # Support alias 'sheet_processing_types' for clarity
        data_sources = self._processing.get('data_sources', self._processing.get('sheet_processing_types', {}))
        return data_sources.get(sheet_name)
    
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
            'data_config': self.get_data_config(sheet_name),
            'context_config': self.get_context_config()
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
            return {
                'columns': sheet_styling['columns'],
                'row_contexts': sheet_styling['row_contexts']
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
        """Get layout configuration for a sheet (headers, blanks, static content, merges)."""
        return self._layout_bundle.get(sheet_name, {})
    
    def get_data_config(self, sheet_name: str) -> Dict[str, Any]:
        """Get data configuration for a sheet (mappings, header_info, etc.)."""
        return self._data_bundle.get(sheet_name, {})
    
    def get_context_config(self) -> Dict[str, Any]:
        """Get global context configuration (replacements, features, extensions)."""
        return self._context
    
    def get_replacement_rules(self) -> Dict[str, Any]:
        """Get text replacement rules from context."""
        return self._context.get('replacements', {})
    
    def get_features(self) -> Dict[str, bool]:
        """Get feature flags."""
        return self.raw_config.get('features', {})
        
    def get_template_json_config(self) -> Optional[Dict[str, Any]]:
        """Get the loaded sibling template JSON config if available."""
        return getattr(self, 'template_json_config', None)
    
    def is_bundled_config(self) -> bool:
        """Check if this is a bundled config (v2.1+)."""
        return self.version.startswith('2.1')
    
    # --- Raw Access (for advanced use cases) ---
    
    def get_raw_config(self) -> Dict[str, Any]:
        """Get the raw config dictionary (avoid using this if possible)."""
        return self.raw_config
