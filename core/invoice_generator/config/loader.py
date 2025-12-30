import json
from typing import Dict, Any, List, Optional


class BundledConfigLoader:
    """
    Loader for the new bundled config format (v2.0+).
    
    Provides clean access to config sections:
    - processing: sheets and data sources
    - layout_bundle: structure, data_flow, content, footer per sheet
    - styling_bundle: styling configuration per sheet
    - features: feature flags
    - defaults: default settings
    """
    
    def __init__(self, config_data: Dict[str, Any]):
        """Initialize with loaded config data."""
        self.config_data = config_data
        self._meta = config_data.get('_meta', {})
        self.processing = config_data.get('processing', {})
        self.layout_bundle = config_data.get('layout_bundle', {})
        self.styling_bundle = config_data.get('styling_bundle', {})
        self.features = config_data.get('features', {})
        self.defaults = config_data.get('defaults', {})
        self.data_preparation_hint = config_data.get('data_preparation_module_hint', {})
    
    @property
    def config_version(self) -> str:
        """Get config version from metadata."""
        return self._meta.get('config_version', 'unknown')
    
    @property
    def customer(self) -> str:
        """Get customer name from metadata."""
        return self._meta.get('customer', '')
    
    # ========== Processing Configuration ==========
    
    def get_sheets_to_process(self) -> List[str]:
        """Get list of sheets to process."""
        return self.processing.get('sheets', [])
    
    def get_data_source(self, sheet_name: str) -> Optional[str]:
        """Get data source for a specific sheet."""
        data_sources = self.processing.get('data_sources', {})
        return data_sources.get(sheet_name)
    
    def get_sheet_data_map(self) -> Dict[str, str]:
        """Get complete sheet to data source mapping."""
        return self.processing.get('data_sources', {})
    
    # ========== Layout Configuration ==========
    
    def get_sheet_layout(self, sheet_name: str) -> Dict[str, Any]:
        """Get complete layout configuration for a sheet."""
        return self.layout_bundle.get(sheet_name, {})
    
    def get_sheet_structure(self, sheet_name: str) -> Dict[str, Any]:
        """Get structure section (start_row, columns) for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('structure', {})
    
    def get_sheet_data_flow(self, sheet_name: str) -> Dict[str, Any]:
        """Get data_flow section (mappings) for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('data_flow', {})
    
    def get_sheet_content(self, sheet_name: str) -> Dict[str, Any]:
        """Get content section (static content) for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('content', {})
    
    def get_sheet_footer(self, sheet_name: str) -> Dict[str, Any]:
        """Get footer section for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('footer', {})
    
    # ========== Styling Configuration ==========
    
    def get_sheet_styling(self, sheet_name: str) -> Dict[str, Any]:
        """Get complete styling configuration for a sheet."""
        return self.styling_bundle.get(sheet_name, {})
    
    def get_styling_defaults(self) -> Dict[str, Any]:
        """Get default styling configuration."""
        return self.styling_bundle.get('defaults', {})
    
    # ========== Features ==========
    
    def is_feature_enabled(self, feature_name: str) -> bool:
        """Check if a feature is enabled."""
        return self.features.get(feature_name, False)


def load_config(config_path: str) -> Dict[str, Any]:
    """Loads the main configuration from a JSON file."""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_bundled_config(config_path: str) -> BundledConfigLoader:
    """Loads and wraps a bundled config file."""
    config_data = load_config(config_path)
    return BundledConfigLoader(config_data)



