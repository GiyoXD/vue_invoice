import logging
from typing import Dict, Any, Optional
from .data_router import get_data_source_for_type
from core.system_config import ConfigurationError

logger = logging.getLogger(__name__)

class BundleResolver:
    """Resolves and extracts static configuration bundles (style, layout, mappings)."""
    
    def __init__(self, sheet_config: Dict[str, Any], sheet_name: str, args: Any = None, invoice_data: Optional[Dict[str, Any]] = None):
        self._sheet_config = sheet_config
        self.sheet_name = sheet_name
        self.args = args
        self.invoice_data = invoice_data

    def get_style_bundle(self) -> Dict[str, Any]:
        return {
            'styling_config': self._sheet_config.get('styling_config', {})
        }

    def get_layout_bundle(self) -> Dict[str, Any]:
        layout_config = self._sheet_config.get('layout_config', {})
        static_payload = layout_config.get('static_payload', {})

        return {
            'sheet_config': layout_config,
            'blanks': layout_config.get('blanks', {}),
            'static_payload': static_payload,
            'merge_rules': layout_config.get('merge_rules', {}),
        }

    def get_data_bundle(self, table_key: Optional[str] = None) -> Dict[str, Any]:
        layout_config = self._sheet_config.get('layout_config', {})
        
        # Resolve active mode from args
        mode = "standard"
        if self.args:
            if getattr(self.args, 'DAF', False): mode = "daf"
            elif getattr(self.args, 'custom', False): mode = "custom"
        
        # Resolve data source type — supports both string and per-mode dict
        raw_source = self._sheet_config.get('data_source', 'aggregation')
        if isinstance(raw_source, dict):
            data_source_type = raw_source.get(mode, raw_source.get('standard', 'standard'))
        elif raw_source == 'aggregation':
            # Classification type → actual data key is the mode name
            data_source_type = mode
        else:
            data_source_type = raw_source
        
        data_source = get_data_source_for_type(data_source_type, self.invoice_data, self.sheet_name)
        
        if table_key and isinstance(data_source, list):
            try:
                idx = int(table_key)
                if 0 <= idx < len(data_source):
                    data_source = data_source[idx]
                else:
                    logger.warning(f"Resolver: table_key '{table_key}' is out of bounds for list data source.")
                    data_source = []
            except ValueError:
                logger.warning(f"Resolver: Invalid table_key '{table_key}' for list data source.")
        elif table_key and isinstance(data_source, dict):
            has_table_keys = any(str(k).isdigit() for k in data_source.keys())
            if has_table_keys:
                data_source = data_source.get(str(table_key), {})
        
        data_flow = layout_config.get('data_flow')
        if not isinstance(data_flow, dict):
            raise ConfigurationError(
                f"CRITICAL: 'data_flow' is missing or not a dictionary in layout configuration for sheet '{self.sheet_name}'."
            )
            
        mapping_rules = data_flow.get('mappings')
        if not isinstance(mapping_rules, dict):
            raise ConfigurationError(
                f"CRITICAL: 'mappings' is missing or not a dictionary in layout_config.data_flow for sheet '{self.sheet_name}'."
            )
        
        return {
            'data_source': data_source,
            'data_source_type': data_source_type,
            'mapping_rules': mapping_rules,
            'static_payload': layout_config.get('static_payload', {}),
            'table_key': table_key,
            'footer_data': self.invoice_data.get('footer_data', {}) if self.invoice_data else {},
        }
