# invoice_generator/processors/single_table_processor.py
import sys
import logging
from .base_processor import SheetProcessor
from ..utils import text_replace_utils
from ..builders.layout_builder import LayoutBuilder
from ..config.builder_config_resolver import BuilderConfigResolver

logger = logging.getLogger(__name__)

class SingleTableProcessor(SheetProcessor):
    """
    Processes a worksheet that is configured to have a single main data table.
    This includes writing a header, filling the table, and applying styles.
    """
    def process(self) -> bool:
        """
        Executes the logic for processing a single-table sheet using the builder pattern.
        """
        logger.info(f"Processing sheet '{self.sheet_name}' as single table/aggregation")
        
        # Calculate weight totals from processed_tables_data (similar to pallet totals)
        from decimal import Decimal, InvalidOperation
        total_net_weight = Decimal('0')
        total_gross_weight = Decimal('0')
        
        if self.invoice_data and 'processed_tables_data' in self.invoice_data:
            processed_tables = self.invoice_data['processed_tables_data']
            # For single table sheets, use first table (usually '1')
            first_table_key = list(processed_tables.keys())[0] if processed_tables else None
            if first_table_key:
                table_data = processed_tables[first_table_key]
                net_weights = table_data.get('net', [])
                gross_weights = table_data.get('gross', [])
                
                for weight in net_weights:
                    try:
                        total_net_weight += Decimal(str(weight))
                    except (InvalidOperation, TypeError, ValueError):
                        continue
                
                for weight in gross_weights:
                    try:
                        total_gross_weight += Decimal(str(weight))
                    except (InvalidOperation, TypeError, ValueError):
                        continue
        
        logger.debug(f"Calculated weight totals for {self.sheet_name}: N.W={total_net_weight}, G.W={total_gross_weight}")
        
        # Use BuilderConfigResolver to prepare bundles cleanly
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name=self.sheet_name,
            worksheet=self.output_worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=self.final_grand_total_pallets,
            final_grand_total_pallets=self.final_grand_total_pallets,  # Context override
            total_net_weight=float(total_net_weight),  # Add weight totals to context
            total_gross_weight=float(total_gross_weight)
        )
        
        # Get the bundles needed for LayoutBuilder
        style_config = resolver.get_style_bundle()
        context_config = resolver.get_context_bundle(
            invoice_data=self.invoice_data,
            enable_text_replacement=False  # Already done at main level
        )
        layout_config = resolver.get_layout_bundle()
        layout_config['enable_text_replacement'] = False
        layout_config['skip_data_table_builder'] = False  # IMPORTANT: Enable data table builder to use resolver
        
        logger.debug(f"layout_config keys: {list(layout_config.keys())}")
        logger.debug(f"skip_data_table_builder in layout_config: {layout_config.get('skip_data_table_builder', 'NOT SET')}")
        logger.debug(f"skip_data_table_builder in sheet_config: {layout_config.get('sheet_config', {}).get('skip_data_table_builder', 'NOT SET')}")
        
        # Get data bundle to extract header_info and mapping_rules
        data_bundle = resolver.get_data_bundle()
        layout_config['header_info'] = data_bundle.get('header_info', {})
        layout_config['mapping_rules'] = data_bundle.get('mapping_rules', {})
        layout_config['data_source'] = data_bundle.get('data_source')
        layout_config['data_source_type'] = data_bundle.get('data_source_type')
        # NOTE: header_info from config is just column metadata, NOT styled Excel rows
        # HeaderBuilder still needs to run to write the actual styled header rows
        
        logger.debug(f"header_info keys: {list(data_bundle.get('header_info', {}).keys())}")
        
        # NEW: Use TableDataAdapter to prepare data
        try:
            table_resolver = resolver.get_table_data_resolver()
            resolved_data = table_resolver.resolve()
            layout_config['resolved_data'] = resolved_data
            logger.info("Successfully resolved table data using TableDataAdapter")
        except Exception as e:
            logger.error(f"Error resolving table data: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        # Use LayoutBuilder to orchestrate the entire layout construction
        layout_builder = LayoutBuilder(
            self.output_workbook,
            self.output_worksheet,
            self.template_worksheet,
            style_config=style_config,
            context_config=context_config,
            layout_config=layout_config,
            template_json_config=self.config_loader.get_template_json_config()
        )
        
        # Build the entire layout (header + table + footer)
        success = layout_builder.build()
        
        if not success:
            logger.error(f"Failed to build layout for sheet '{self.sheet_name}'")
            return False
            
        logger.info(f"Successfully filled table data/footer for sheet '{self.sheet_name}'")
        
        # TODO: Re-implement post-processing features using new architecture:
        # - Weight summary (should be a builder add-on)
        # - Column widths (should be handled by styling in builders)
        # - Summary fields (should be part of data mapping)
        
        return True
