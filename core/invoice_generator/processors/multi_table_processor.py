# invoice_generator/processors/multi_table_processor.py
import sys
import logging
import traceback
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter

from .base_processor import SheetProcessor
from ..builders.layout_builder import LayoutBuilder
from ..builders.table import TableFooterBuilder
from ..builders.table.table_grid import Grid
from ..builders.json_template_builder import JsonTemplateStateBuilder
from ..styling.models import StylingConfigModel
from ..mappers import resolve_flat_footer_payload
from ..models.footer import FooterData
from ..models.layout import SheetLayoutState, TableLayoutConfig
from ..models.config.layout import FooterConfigModel
from ..resolvers.sheet_config_resolver import SheetConfigResolver
from ..utils.cell_converter import write_models_to_worksheet
from core.system_config import ConfigurationError

logger = logging.getLogger(__name__)

class MultiTableProcessor(SheetProcessor):
    """
    Processes a worksheet that contains multiple, repeating blocks of tables,
    such as a packing list. Uses LayoutBuilder for each table iteration.
    """

    def process(self) -> bool:
        """
        Executes the logic for processing a multi-table sheet using LayoutBuilder.
        """
        logger.info(f"Processing sheet '{self.sheet_name}' as multi-table/packing list")
        
        # 1. Resolve Data
        all_tables_data, table_keys = self._resolve_all_tables_data()
        if not all_tables_data:
            return True  # Nothing to do

        # Evaluate global col_desc merge rules
        # self.is_global_unique_desc is already calculated by BaseProcessor
        self.allow_col_desc_merge = True # Authority flag to enable/disable the feature entirely
        
        logger.info(f"MultiTableProcessor: Global Unique Description Flag: {self.is_global_unique_desc}")

        # 2. Capture Template State
        template_state_builder = self._capture_template_state()
        if not template_state_builder:
            return False

        # 3. Initialize Tracking Variables
        layout_state = SheetLayoutState()
        layout_state.advance_to(self.header_row)
        
        current_row = self.header_row
        all_data_ranges = []
        grand_total_pallets = 0
        last_grid = None
        
        # 4. Process Each Table
        for i, table_key in enumerate(table_keys):
            is_first_table = (i == 0)
            is_last_table = (i == len(table_keys) - 1)
            
            logger.info(f"Processing table '{table_key}' ({i+1}/{len(table_keys)})")
            
            layout_builder = self._build_table_layout(
                layout_state=layout_state,
                table_key=table_key,
                config=TableLayoutConfig(
                    is_first_table=is_first_table,
                    is_last_table=is_last_table,
                    skip_template_footer=True
                ),
                template_state_builder=template_state_builder
            )
            
            if not layout_builder:
                return False
            
            # Calculate next row
            next_row = layout_builder.next_row_after_footer
            if not is_last_table:
                next_row += 1
                
            # Retrieve pallet count from LayoutBuilder (calculated by TableCalculator)
            table_pallets = layout_builder.footer_data.total_pallets if layout_builder.footer_data else 0
            
            # Get data range
            data_range = None
            if layout_builder.data_start_row > 0 and layout_builder.data_end_row >= layout_builder.data_start_row:
                data_range = (layout_builder.data_start_row, layout_builder.data_end_row)
            
            # Update tracking
            layout_state.advance_to(next_row)
            current_row = layout_state.next_free_row
            grand_total_pallets += table_pallets
            if data_range:
                all_data_ranges.append(data_range)
            last_grid = layout_builder.grid

        # 5. Build Grand Total Row
        if len(table_keys) > 1 and last_grid:
            current_row = self._build_grand_total_row(
                current_row=current_row,
                all_data_ranges=all_data_ranges,
                last_grid=last_grid
            )

        # 6. Build Page-level Summary
        if last_grid and layout_builder:
            current_row = self._build_page_summary(
                grid=last_grid,
                sheet_layout=layout_builder.sheet_layout,
                footer_data=layout_builder.footer_data
            )

        # 7. Restore Template Footer
        self._restore_template_footer(template_state_builder, current_row, last_grid)
        
        logger.info(f"Successfully processed {len(table_keys)} tables for sheet '{self.sheet_name}'.")
        return True

    def _resolve_all_tables_data(self) -> Tuple[Optional[List], List]:
        """Resolves all tables data using SheetConfigResolver."""
        initial_resolver = SheetConfigResolver(
            config_loader=self.config_loader,
            sheet_name=self.sheet_name,
            worksheet=self.output_worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=0
        )
        
        all_tables_data = initial_resolver._get_data_source_for_type('processed_tables_multi')
        if not all_tables_data or not isinstance(all_tables_data, list):
            logger.warning(f"'processed_tables_data' not found/valid or is not a list. Skipping '{self.sheet_name}'")
            return None, []

        table_keys = [str(i) for i in range(len(all_tables_data))]
        logger.info(f"Found {len(table_keys)} tables to process")
        return all_tables_data, table_keys

    def _capture_template_state(self):
        """Captures template state (header/footer) for reuse."""
        logger.info(f"[MultiTableProcessor] Capturing template state for reuse")
        
        # Check for JSON config - this is now REQUIRED
        json_config = self.config_loader.get_template_json_config()
        if json_config and self.sheet_name in json_config:
            logger.info(f"Using JSON-based template state")
            try:
                sheet_layout_json = json_config[self.sheet_name]
                template_state_builder = JsonTemplateStateBuilder(
                    sheet_layout_data=sheet_layout_json
                )
                return template_state_builder
            except Exception as e:
                logger.critical(f"CRITICAL: JsonTemplateStateBuilder failed: {e}", exc_info=True)
                return None

        # JSON template is required - XLSX scanning has been removed
        logger.critical(f"CRITICAL: No JSON template found for sheet '{self.sheet_name}'. XLSX scanning has been removed.")
        return None

    def _build_grand_total_row(self, current_row, all_data_ranges, last_grid):
        """Builds the Grand Total row after all tables."""
        logger.info("Adding Grand Total Row")
        
        # Prepare footer config
        raw_footer = self.layout_config.get('footer', {}) if self.layout_config else {}
        footer_config = FooterConfigModel.model_validate(raw_footer) if raw_footer else FooterConfigModel()
        footer_config.type = "grand_total"
        
        # Build the payload
        footer_dict = self.invoice_data.get('footer_data', {}) if self.invoice_data else {}
        raw_gt = footer_dict.get('grand_total', {})
        gt_payload = resolve_flat_footer_payload(raw_gt)
        gt_payload['leather_summary'] = footer_dict.get('leather_summary', [])

        # Reuse last_grid's registries to build gt_grid without duplicating Registry/Styling setups
        gt_grid = Grid(
            column_mapping=last_grid.column_mapping,
            style_registry=last_grid.style_registry,
            column_colspan=last_grid.column_colspan,
            dimension_registry=last_grid.dimension_registry
        )
        gt_grid.set_start_row(current_row)

        footer_builder = TableFooterBuilder(
            grid=gt_grid,
            footer_config=footer_config,
            payload=gt_payload,
            sum_ranges=all_data_ranges
        )
        
        try:
            footer_builder.build()
        except Exception as e:
            logger.error("Failed to build grand total footer", exc_info=True)
            return current_row
            
        row_models = gt_grid.get_row_models()
        if row_models:
            write_models_to_worksheet(self.output_worksheet, row_models, start_row=current_row)
            
        return current_row + gt_grid._cursor_row
