# invoice_generator/config/builder_config_resolver.py
import logging
from typing import Any, Dict, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet

from .resolvers import BundleResolver, ContextResolver, get_data_source_for_type

logger = logging.getLogger(__name__)


class BuilderConfigResolver:
    """
    Resolves and prepares configuration bundles for specific builders by
    delegating to the dedicated resolvers in the config/resolvers/ subpackage.
    """
    
    def __init__(
        self,
        config_loader,
        sheet_name: str,
        worksheet: Worksheet,
        args=None,
        invoice_data: Optional[Dict[str, Any]] = None,
        pallets: int = 0,
        **context_overrides
    ):
        self.config_loader = config_loader
        self.sheet_name = sheet_name
        self.worksheet = worksheet
        self.args = args
        self.invoice_data = invoice_data
        self.pallets = pallets
        self.context_overrides = context_overrides
        
        # Instantiate sub-resolvers
        self._sheet_config = config_loader.get_sheet_config(sheet_name)
        self.bundle_resolver = BundleResolver(self._sheet_config, sheet_name, args, invoice_data)
        self.context_resolver = ContextResolver(sheet_name, args, invoice_data, pallets, config_loader, context_overrides)
        
    def get_style_bundle(self) -> Dict[str, Any]:
        return self.bundle_resolver.get_style_bundle()
        
    def get_context_bundle(self, table_key: Optional[str] = None, **additional_context) -> Dict[str, Any]:
        return self.context_resolver.get_context_bundle(table_key=table_key, **additional_context)
        
    def get_layout_bundle(self) -> Dict[str, Any]:
        return self.bundle_resolver.get_layout_bundle()
        
    def get_data_bundle(self, table_key: Optional[str] = None) -> Dict[str, Any]:
        return self.bundle_resolver.get_data_bundle(table_key=table_key)
        
    def get_header_bundles(self) -> Tuple[Dict, Dict, Dict]:
        return self.get_style_bundle(), self.get_context_bundle(), self.get_layout_bundle()
        
    def get_datatable_bundles(self, table_key: Optional[str] = None) -> Tuple[Dict, Dict, Dict, Dict]:
        return self.get_style_bundle(), self.get_context_bundle(), self.get_layout_bundle(), self.get_data_bundle(table_key=table_key)
        
    def get_layout_bundles_with_data(self, table_key: Optional[str] = None) -> Tuple[Dict, Dict, Dict]:
        style_config = self.get_style_bundle()
        context_config = self.get_context_bundle(table_key=table_key)
        layout_config = self.get_layout_bundle()
        data_config = self.get_data_bundle(table_key=table_key)
        
        merged_layout_config = {
            **layout_config,
            'data_source': data_config.get('data_source'),
            'data_source_type': data_config.get('data_source_type'),
            'mapping_rules': data_config.get('mapping_rules'),
        }
        return style_config, context_config, merged_layout_config
        
    def get_table_data_resolver(self, table_key: Optional[str] = None):
        from .table_value_adapter import TableDataAdapter
        return TableDataAdapter.create_from_bundles(
            data_config=self.get_data_bundle(table_key=table_key),
            context_config=self.get_context_bundle(),
            layout_config=self.get_layout_bundle()
        )
        
    def get_table_footer_resolver(self, table_key: Optional[str] = None):
        from .table_value_adapter import TableFooterAdapter
        return TableFooterAdapter.create_from_bundles(
            data_config=self.get_data_bundle(table_key=table_key),
            context_config=self.get_context_bundle()
        )
        
    def get_footer_bundles(
        self,
        sum_ranges: Optional[list] = None,
        pallet_count: Optional[int] = None,
        is_last_table: bool = False
    ) -> Tuple[Dict, Dict, Dict]:
        style_config = self.get_style_bundle()
        context_config = self.get_context_bundle(
            pallet_count=pallet_count if pallet_count is not None else self.pallets,
            is_last_table=is_last_table
        )
        data_config = self.get_data_bundle()
        data_config.update({
            'sum_ranges': sum_ranges or [],
            'footer_config': self._sheet_config.get('layout_config', {}).get('footer', {}),
            'DAF_mode': self.args.DAF if self.args and hasattr(self.args, 'DAF') else False,
            'custom_mode': self.args.custom if self.args and hasattr(self.args, 'custom') else False,
        })
        return style_config, context_config, data_config
        
    def get_footer_data(
        self,
        footer_row_start_idx: int,
        data_start_row: int,
        data_end_row: int,
        pallet_count: Optional[int] = None,
        leather_summary: Optional[Dict] = None,
        weight_summary: Optional[Dict] = None
    ):
        return self.context_resolver.get_footer_data(
            footer_row_start_idx=footer_row_start_idx,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            pallet_count=pallet_count,
            leather_summary=leather_summary,
            weight_summary=weight_summary
        )
        
    def get_all_sheet_configs(self) -> Dict[str, Any]:
        return self.config_loader.get_raw_config().get('layout_bundle', {})
        
    def _get_data_source_for_type(self, data_source_type: str) -> Any:
        return get_data_source_for_type(
            data_source_type=data_source_type,
            invoice_data=self.invoice_data,
            sheet_name=self.sheet_name,
            args=self.args
        )
