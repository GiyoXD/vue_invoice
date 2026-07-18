# invoice_generator/processors/single_table_processor.py
import sys
import logging
from .base_processor import SheetProcessor
from ..builders.layout_builder import LayoutBuilder
from ..resolvers.sheet_config_resolver import SheetConfigResolver

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
        # Retrieve pre-calculated total weights from footer_data.grand_total
        footer_data = self.invoice_data.get('footer_data', {}) if self.invoice_data else {}
        grand_total = footer_data.get('grand_total', {})
        total_net_weight_val = grand_total.get('col_net')
        total_gross_weight_val = grand_total.get('col_gross')

        total_net_weight = float(total_net_weight_val) if total_net_weight_val is not None else 0.0
        total_gross_weight = float(total_gross_weight_val) if total_gross_weight_val is not None else 0.0
        
        logger.debug(f"Resolved weight totals for {self.sheet_name}: N.W={total_net_weight}, G.W={total_gross_weight}")
        
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
