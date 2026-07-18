# invoice_generator/processors/base_processor.py
from abc import ABC, abstractmethod
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import argparse
from typing import Dict, Any, Optional
from core.system_config import ConfigurationError
from core.invoice_generator.models.context import ProcessorContext
from core.invoice_generator.models.layout import SheetLayoutState, TableLayoutConfig

class SheetProcessor(ABC):
    """
    Abstract base class for processing a single worksheet in an invoice workbook.
    Defines the common interface for all concrete processor implementations.
    """
    def __init__(self, ctx: ProcessorContext):
        """
        Initializes the processor with all necessary data and configurations.

        Args:
            ctx: The structured ProcessorContext holding io, config, and data sub-contexts.
        """
        self.template_workbook = ctx.io.template_workbook
        self.output_workbook = ctx.io.output_workbook
        self.template_worksheet = ctx.io.template_worksheet
        self.output_worksheet = ctx.io.output_worksheet
        
        # Keep old names for backward compatibility during transition
        self.workbook = ctx.io.output_workbook
        self.worksheet = ctx.io.output_worksheet
        self.output_worksheet = ctx.io.output_worksheet
        
        self.sheet_name = ctx.config.sheet_name
        self.sheet_config = ctx.config.sheet_config
        self.data_source_indicator = ctx.config.data_source_indicator
        self.config_loader = ctx.config.config_loader
        
        self.invoice_data = ctx.data.invoice_data
        self.args = ctx.data.cli_args
        self.processing_successful = True
        self._use_bundled = self.config_loader is not None
        self.data_mapping_config = None  # Deprecated
        
        # New: Strict Header Row Validation (Mandatory for all table-based sheets)
        self.layout_config = self.sheet_config.get('layout_config', {}) if self.sheet_config else {}
        structure = self.layout_config.get('structure', {})
        self.header_row = structure.get('header_row')
        
        if self.header_row is None:
             raise ConfigurationError(f"CRITICAL: No 'header_row' found in configuration for sheet '{self.sheet_name}'. "
                                    f"Please ensure layout_config -> structure -> header_row is defined in your JSON.")

        # --- GLOBAL UNIQUENESS SCAN (Invoice Scope) ---
        # Calculate once in BaseProcessor so all children share the same authority.
        self.all_global_descriptions = set()
        
        # 1. Scan raw data
        if self.invoice_data:
            # Check multi_table
            multi_table = self.invoice_data.get('multi_table', [])
            if isinstance(multi_table, list):
                for table in multi_table:
                    if isinstance(table, list):
                        for row in table:
                            d = str(row.get('col_desc', "")).strip()
                            if d: self.all_global_descriptions.add(d)
            
            # Check single_table
            single_table = self.invoice_data.get('single_table', {})
            if isinstance(single_table, dict):
                for agg_key in ['aggregation', 'aggregation_custom', 'aggregation_DAF']:
                    agg_data = single_table.get(agg_key, [])
                    if isinstance(agg_data, list):
                        for row in agg_data:
                            d = str(row.get('col_desc', "")).strip()
                            if d: self.all_global_descriptions.add(d)

        # 2. Scan Fallbacks in Configuration (Truth if data is empty)
        if self.config_loader:
            raw_config = self.config_loader.get_raw_config()
            layout_bundle = raw_config.get('layout_bundle', {})
            for sheet_name, sheet_conf in layout_bundle.items():
                if not isinstance(sheet_conf, dict):
                    continue
                    
                # Check data_flow -> mappings -> col_desc
                mappings = sheet_conf.get('data_flow', {}).get('mappings', {})
                col_desc_rule = mappings.get('col_desc', {})
                if isinstance(col_desc_rule, dict):
                    fallback = col_desc_rule.get('fallback')
                    if isinstance(fallback, dict):
                        # Modern nested format
                        for mode_val in fallback.values():
                            if isinstance(mode_val, str) and mode_val.strip():
                                self.all_global_descriptions.add(mode_val.strip())
                    elif isinstance(fallback, str) and fallback.strip():
                        self.all_global_descriptions.add(fallback.strip())
        
        # Determine global uniqueness flag
        self.is_global_unique_desc = (len(self.all_global_descriptions) <= 1)
        # ---------------------------------------------

    @abstractmethod
    def process(self) -> bool:
        """
        Main method to orchestrate the processing of the worksheet.
        This must be implemented by all subclasses.

        Returns:
            bool: True if processing was successful, False otherwise.
        """
        pass

    def _build_table_layout(
        self,
        layout_state: SheetLayoutState,
        table_key: Optional[str],
        config: TableLayoutConfig,
        template_state_builder: Optional[Any] = None
    ):
        """
        Orchestrates the common workflow of SheetConfigResolver resolution and LayoutBuilder execution.
        Saves subclasses from duplicating this execution sequence.
        """
        resolver = self._init_resolver(
            is_last_table=config.is_last_table,
            total_net_weight=config.total_net_weight,
            total_gross_weight=config.total_gross_weight
        )
        
        style_config, context_config, layout_config = self._resolve_layout_configs(
            resolver=resolver,
            table_key=table_key,
            is_last_table=config.is_last_table,
            next_free_row=layout_state.next_free_row
        )
        
        if not self._resolve_table_data(resolver, table_key, layout_config):
            return None
            
        return self._run_layout_builder(
            layout_state=layout_state,
            style_config=style_config,
            context_config=context_config,
            layout_config=layout_config,
            template_state_builder=template_state_builder,
            is_first_table=config.is_first_table,
            skip_template_footer=config.skip_template_footer
        )

    def _init_resolver(
        self,
        is_last_table: bool,
        total_net_weight: Optional[float],
        total_gross_weight: Optional[float]
    ):
        """Initializes the SheetConfigResolver with context overrides."""
        from core.invoice_generator.resolvers.sheet_config_resolver import SheetConfigResolver

        context_overrides = {}
        if total_net_weight is not None:
            context_overrides["total_net_weight"] = total_net_weight
        if total_gross_weight is not None:
            context_overrides["total_gross_weight"] = total_gross_weight
        
        # Add context override
        context_overrides["is_last_table"] = is_last_table

        return SheetConfigResolver(
            config_loader=self.config_loader,
            sheet_name=self.sheet_name,
            worksheet=self.output_worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=0,
            **context_overrides
        )

    def _resolve_layout_configs(
        self,
        resolver,
        table_key: Optional[str],
        is_last_table: bool,
        next_free_row: int
    ):
        """Resolves style, context, and layout config bundles."""
        style_config = resolver.get_style_bundle()
        context_config = resolver.get_context_bundle(
            table_key=table_key,
            is_last_table=is_last_table
        )
        layout_config = resolver.get_layout_bundle()
        
        # Override header row position for legacy compatibility
        if not 'structure' in layout_config.get('sheet_config', {}):
            if 'sheet_config' not in layout_config:
                layout_config['sheet_config'] = {}
            layout_config['sheet_config']['structure'] = {}
        layout_config['sheet_config']['structure']['header_row'] = next_free_row
        
        # Enable data table builder when processing a single table (meaning table_key is None)
        if table_key is None:
            layout_config['skip_data_table_builder'] = False

        return style_config, context_config, layout_config

    def _resolve_table_data(self, resolver, table_key: Optional[str], layout_config: Dict[str, Any]) -> bool:
        """Resolves table data using TableDataMapper and updates layout_config."""
        import logging
        logger = logging.getLogger(__name__)

        data_bundle = resolver.get_data_bundle(table_key=table_key)
        layout_config['mapping_rules'] = data_bundle.get('mapping_rules', {})
        layout_config['data_source'] = data_bundle.get('data_source')
        layout_config['data_source_type'] = data_bundle.get('data_source_type')
        
        try:
            table_resolver = resolver.get_table_data_resolver(table_key=table_key)
            resolved_data = table_resolver.resolve()
            layout_config['resolved_data'] = resolved_data
            
            logger.info(f"Successfully resolved table data for table '{table_key or 'default'}' using TableDataMapper")
            return True
        except Exception as e:
            logger.error(f"Error resolving table data: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _run_layout_builder(
        self,
        layout_state,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Dict[str, Any],
        template_state_builder: Optional[Any],
        is_first_table: bool,
        skip_template_footer: bool
    ):
        """Runs the LayoutBuilder and returns layout_builder if successful."""
        import logging
        logger = logging.getLogger(__name__)
        from core.invoice_generator.builders.layout_builder import LayoutBuilder
        from core.invoice_generator.models.config.styling import SheetStylingModel
        from core.invoice_generator.models.config.layout import SheetLayoutModel

        layout_config['skip_template_header_restoration'] = (not is_first_table)
        layout_config['skip_template_footer_restoration'] = skip_template_footer
        layout_config['allow_col_desc_merge'] = getattr(self, 'allow_col_desc_merge', True)
        layout_config['is_global_unique_desc'] = getattr(self, 'is_global_unique_desc', False)

        # Set allow_col_desc_merge and is_global_unique_desc on self.args so they are propagated downstream
        if self.args:
            self.args.allow_col_desc_merge = layout_config['allow_col_desc_merge']
            self.args.is_global_unique_desc = layout_config['is_global_unique_desc']
        
        sheet_styling = SheetStylingModel.model_validate(style_config.get('styling_config', {}))
        sheet_layout = SheetLayoutModel.model_validate(layout_config.get('sheet_config', {}))
        resolved_data = layout_config.get('resolved_data')
        
        layout_builder = LayoutBuilder(
            workbook=self.output_workbook,
            worksheet=self.output_worksheet,
            template_worksheet=self.template_worksheet,
            sheet_styling=sheet_styling,
            sheet_layout=sheet_layout,
            resolved_data=resolved_data,
            sheet_name=self.sheet_name,
            all_sheet_configs=context_config.get('all_sheet_configs', {}),
            args=self.args,
            total_net_weight=context_config.get('total_net_weight'),
            total_gross_weight=context_config.get('total_gross_weight'),
            is_last_table=context_config.get('is_last_table', False),
            skip_template_header_restoration=layout_config.get('skip_template_header_restoration', False),
            skip_header_builder=layout_config.get('skip_header_builder', False),
            skip_data_table_builder=layout_config.get('skip_data_table_builder', False),
            skip_footer_builder=layout_config.get('skip_footer_builder', False),
            skip_template_footer_restoration=layout_config.get('skip_template_footer_restoration', False),
            template_state_builder=template_state_builder,
            template_json_config=self.config_loader.get_template_json_config(),
            layout_state=layout_state
        )
        
        success = layout_builder.build()
        if not success:
            logger.error("Failed to build layout")
            return None
            
        return layout_builder

