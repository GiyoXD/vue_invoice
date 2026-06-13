# invoice_generator/processors/single_table_processor.py
import sys
import logging
from .base_processor import SheetProcessor
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
        
        from core.invoice_generator.models.layout import SheetLayoutState, TableLayoutConfig
        layout_state = SheetLayoutState()
        layout_state.advance_to(self.header_row)

        layout_builder = self._build_table_layout(
            layout_state=layout_state,
            table_key=None,
            config=TableLayoutConfig(
                is_first_table=True,
                is_last_table=True,
                total_net_weight=float(total_net_weight),
                total_gross_weight=float(total_gross_weight)
            )
        )
        
        if not layout_builder:
            logger.error(f"Failed to build layout for sheet '{self.sheet_name}'")
            return False
            
        logger.info(f"Successfully filled table data/footer for sheet '{self.sheet_name}'")
        return True
