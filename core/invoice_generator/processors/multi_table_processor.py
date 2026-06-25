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
from ..styling.models import StylingConfigModel
from ..models.footer import FooterData
from ..config.builder_config_resolver import BuilderConfigResolver

logger = logging.getLogger(__name__)
from core.system_config import ConfigurationError

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
        from core.invoice_generator.models.layout import SheetLayoutState
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
            show_grand_total_addons = (len(table_keys) == 1)
            
            logger.info(f"Processing table '{table_key}' ({i+1}/{len(table_keys)})")
            
            from core.invoice_generator.models.layout import TableLayoutConfig
            layout_builder = self._build_table_layout(
                layout_state=layout_state,
                table_key=table_key,
                config=TableLayoutConfig(
                    is_first_table=is_first_table,
                    is_last_table=is_last_table,
                    skip_template_footer=True,
                    show_grand_total_addons=show_grand_total_addons
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
                grand_total_pallets=grand_total_pallets,
                all_data_ranges=all_data_ranges,
                last_grid=last_grid,
                all_tables_data=all_tables_data,
                table_keys=table_keys
            )

        # 6. Restore Template Footer
        self._restore_template_footer(template_state_builder, current_row, table_keys, last_grid)
        
        logger.info(f"Successfully processed {len(table_keys)} tables for sheet '{self.sheet_name}'.")
        return True

    def _resolve_all_tables_data(self) -> Tuple[Optional[List], List]:
        """Resolves all tables data using BuilderConfigResolver."""
        initial_resolver = BuilderConfigResolver(
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
        from ..builders.json_template_builder import JsonTemplateStateBuilder
        
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

    def _build_grand_total_row(self, current_row, grand_total_pallets, all_data_ranges, last_grid, 
                             all_tables_data, table_keys):
        """Builds the Grand Total row after all tables."""
        logger.info("Adding Grand Total Row")
        
        # Fetch the global leather summary securely to prevent data multiplication
        global_leather_summary = {}
        if self.invoice_data and 'footer_data' in self.invoice_data:
            footer_data = self.invoice_data.get('footer_data', {})
            add_ons = footer_data.get('add_ons', {})
            if add_ons:
                global_leather_summary = add_ons.get('leather_summary_addon', {})
        
        grand_total_resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name=self.sheet_name,
            worksheet=self.output_worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=grand_total_pallets
        )
        
        gt_style_config = grand_total_resolver.get_style_bundle()
        gt_layout_config = grand_total_resolver.get_layout_bundle()
        
        from core.invoice_generator.models.config.styling import SheetStylingModel
        from core.invoice_generator.models.config.layout import SheetLayoutModel, FooterConfigModel
        from ..styling.style_registry import StyleRegistry
        from ..styling.dimension_registry import DimensionRegistry
        
        sheet_styling = SheetStylingModel.model_validate(gt_style_config.get('styling_config', {}))
        sheet_layout = SheetLayoutModel.model_validate(gt_layout_config.get('sheet_config', {}))

        style_registry = StyleRegistry(sheet_styling)
        row_heights = {
            context: style.row_height
            for context, style in sheet_styling.row_contexts.items()
            if style.row_height is not None
        }
        dimension_registry = DimensionRegistry(row_heights)

        from ..builders.table.table_grid import Grid
        gt_grid = Grid(
            column_mapping=last_grid.column_mapping,
            style_registry=style_registry,
            column_colspan=last_grid.column_colspan,
            dimension_registry=dimension_registry
        )
        gt_grid.set_start_row(current_row)

        # Prepare footer config
        footer_config = sheet_layout.footer.model_copy() if sheet_layout.footer else FooterConfigModel()
        footer_config.type = "grand_total"
        
        # Calculate overall data range
        if all_data_ranges:
            overall_data_start = min(r[0] for r in all_data_ranges)
            overall_data_end = max(r[1] for r in all_data_ranges)
        else:
            overall_data_start = current_row - 1
            overall_data_end = current_row - 1
            
        # Create FooterData using resolver to ensure normalized data (including global weights)
        footer_data = grand_total_resolver.get_footer_data(
            footer_row_start_idx=current_row,
            data_start_row=overall_data_start,
            data_end_row=overall_data_end,
            pallet_count=grand_total_pallets,
            leather_summary=global_leather_summary,
            weight_summary={'net': 0.0, 'gross': 0.0}  # Will be auto-filled with global weights by resolver
        )
        
        footer_builder = TableFooterBuilder(
            grid=gt_grid,
            footer_data=footer_data,
            footer_config=footer_config,
            pallet_count=grand_total_pallets,
            show_grand_total_addons=True,
            is_daf=bool(getattr(self.args, 'DAF', False)) if self.args else False,
            sheet_name=self.sheet_name,
            sum_ranges=all_data_ranges
        )
        
        try:
            footer_builder.build()
        except Exception as e:
            logger.error("Failed to build grand total footer", exc_info=True)
            return current_row
            
        row_models = gt_grid.get_row_models()
        from ..utils.cell_converter import write_models_to_worksheet
        if row_models:
            write_models_to_worksheet(self.output_worksheet, row_models, start_row=current_row)
            
        return current_row + gt_grid._cursor_row

    def _restore_template_footer(self, template_state_builder, current_row, table_keys, last_grid):
        """
        Template footer restoration.
        Runs AFTER all tables and the Grand Total row are complete to place 
        static blueprint elements (signatures, warning text) beneath everything.
        """
        if template_state_builder and not self.sheet_config.get('skip_template_footer', False):
            try:
                # We need actual_num_cols which we can get from the last grid
                actual_num_cols = last_grid.num_columns if last_grid else None
                column_index_mapping = getattr(last_grid, 'column_index_mapping', None) if last_grid else None
                
                logger.info(f"--- RESTORING TEMPLATE FOOTER (Multi-Table End) ---")
                logger.info(f"footer_start_row: {current_row}")
                
                # Resolve generation mode for mode-dependent footer values
                gen_mode = "standard"
                if self.args:
                    if getattr(self.args, 'DAF', False): gen_mode = "daf"
                    elif getattr(self.args, 'custom', False): gen_mode = "custom"
 
                template_state_builder.restore_template_footer(
                    target_worksheet=self.output_worksheet,
                    footer_start_row=current_row,
                    actual_num_cols=actual_num_cols,
                    mode=gen_mode,
                    column_index_mapping=column_index_mapping
                )
                logger.info("Template footer restored successfully")
            except Exception as e:
                logger.error(f"Failed to restore template footer: {e}", exc_info=True)
        else:
            logger.debug("Skipping template footer restoration (missing builder or config skip)")
