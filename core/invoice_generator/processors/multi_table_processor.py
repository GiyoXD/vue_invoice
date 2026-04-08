# invoice_generator/processors/multi_table_processor.py
import sys
import logging
import traceback
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter

from .base_processor import SheetProcessor
from ..builders.layout_builder import LayoutBuilder
from ..builders.footer_builder import TableFooterBuilder
from ..styling.models import StylingConfigModel, FooterData
from ..config.builder_config_resolver import BuilderConfigResolver

from ..extractors.header_extractor import HeaderExtractor

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
        # If any table contains varying descriptions in col_desc, disable merge for all tables
        self.allow_col_desc_merge = True
        for table_data in all_tables_data:
            if isinstance(table_data, dict):
                desc_list = table_data.get('col_desc') or table_data.get('desc') or table_data.get('description') or []
                if isinstance(desc_list, list) and len(desc_list) > 1:
                    valid_descs = [str(x).strip() for x in desc_list if x is not None and str(x).strip()]
                    if len(set(valid_descs)) > 1:
                        logger.info(f"MultiTableProcessor: Found varying descriptions in table. Disabling col_desc merge globally.")
                        self.allow_col_desc_merge = False
                        break

        # 2. Capture Template State
        template_state_builder = self._capture_template_state()
        if not template_state_builder:
            return False

        # 3. Initialize Tracking Variables
        current_row = self.header_row
        all_data_ranges = []
        grand_total_pallets = 0
        last_header_info = None
        
        # 4. Process Each Table
        for i, table_key in enumerate(table_keys):
            result = self._process_single_table(
                table_key=table_key,
                index=i,
                total_tables=len(table_keys),
                current_row=current_row,
                all_tables_data=all_tables_data,
                template_state_builder=template_state_builder
            )
            
            if not result:
                return False
            
            # Unpack result
            (next_row, table_pallets, data_range, header_info, table_leather_summary) = result
            
            # Update tracking
            current_row = next_row
            grand_total_pallets += table_pallets
            if data_range:
                all_data_ranges.append(data_range)
            last_header_info = header_info

        # 5. Build Grand Total Row
        if len(table_keys) > 1 and last_header_info:
            current_row = self._build_grand_total_row(
                current_row=current_row,
                grand_total_pallets=grand_total_pallets,
                all_data_ranges=all_data_ranges,
                last_header_info=last_header_info,
                all_tables_data=all_tables_data,
                table_keys=table_keys
            )

        # 6. Restore Template Footer
        self._restore_template_footer(template_state_builder, current_row, table_keys)
        
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
                    sheet_layout_data=sheet_layout_json,
                    debug=getattr(self.args, 'debug', False)
                )
                
                # Extract header info
                if self.args and self.invoice_data:
                    self.header_info = HeaderExtractor.extract(template_state_builder.header_state)
                    
                return template_state_builder
            except Exception as e:
                logger.critical(f"CRITICAL: JsonTemplateStateBuilder failed: {e}", exc_info=True)
                return None

        # JSON template is required - XLSX scanning has been removed
        logger.critical(f"CRITICAL: No JSON template found for sheet '{self.sheet_name}'. XLSX scanning has been removed.")
        return None

    def _process_single_table(self, table_key, index, total_tables, current_row, all_tables_data, template_state_builder):
        """Processes a single table iteration."""
        is_first_table = (index == 0)
        is_last_table = (index == total_tables - 1)
        logger.info(f"Processing table '{table_key}' ({index+1}/{total_tables})")
        
        show_grand_total_addons = (total_tables == 1)
        
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name=self.sheet_name,
            worksheet=self.output_worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=0
        )
        
        style_config = resolver.get_style_bundle()
        context_config = resolver.get_context_bundle(
            is_last_table=is_last_table,
            show_grand_total_addons=show_grand_total_addons
        )
        layout_config = resolver.get_layout_bundle()
        
        # Resolve table data
        table_data_resolver = resolver.get_table_data_resolver(table_key=str(table_key))
        resolved_data = table_data_resolver.resolve()
        layout_config['resolved_data'] = resolved_data
        
        # Override header row position
        if not 'structure' in layout_config.get('sheet_config', {}):
            if 'sheet_config' not in layout_config:
                layout_config['sheet_config'] = {}
            layout_config['sheet_config']['structure'] = {}
        layout_config['sheet_config']['structure']['header_row'] = current_row
        
        layout_config['skip_template_header_restoration'] = (not is_first_table)
        layout_config['skip_template_footer_restoration'] = True
        layout_config['allow_col_desc_merge'] = getattr(self, 'allow_col_desc_merge', True)
        layout_config['data_source_type'] = self.sheet_config.get('data_source', 'processed_tables_multi') if self.sheet_config else 'processed_tables_multi'
        
        layout_builder = LayoutBuilder(
            self.output_workbook,
            self.output_worksheet,
            self.template_worksheet,
            style_config=style_config,
            context_config=context_config,
            layout_config=layout_config,
            template_state_builder=template_state_builder
        )
        
        success = layout_builder.build()
        if not success:
            logger.error(f"Failed to build layout for table '{table_key}'")
            return None
        
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
            
        return (
            next_row,
            table_pallets,
            data_range,
            layout_builder.header_info,
            getattr(layout_builder, 'leather_summary', None)
        )

    def _build_grand_total_row(self, current_row, grand_total_pallets, all_data_ranges, last_header_info, 
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
        
        # Prepare styling model
        styling_model = gt_style_config.get('styling_config')
        if styling_model and not isinstance(styling_model, StylingConfigModel):
            if isinstance(styling_model, dict) and 'columns' in styling_model and 'row_contexts' in styling_model:
                pass
            else:
                try:
                    styling_model = StylingConfigModel(**styling_model)
                except Exception as e:
                    logger.warning(f"Could not create StylingConfigModel: {e}")
                    styling_model = None
        
        # Prepare footer config
        sheet_config = gt_layout_config.get('sheet_config', {})
        footer_config = sheet_config.get('footer', {}).copy()
        footer_config["type"] = "grand_total"
        
        # Legacy DAF summary logic removed - add_ons are now controlled via dictionary config
        # if sheet_config.get('content', {}).get("summary", False) and self.args.DAF:
        #     footer_config["add_ons"] = ["summary"]
        
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
            worksheet=self.output_worksheet,
            footer_data=footer_data,
            style_config={'styling_config': styling_model},
            context_config={
                'header_info': last_header_info,
                'pallet_count': grand_total_pallets,
                'sheet_name': self.sheet_name,
                'is_last_table': True
            },
            data_config={
                'sum_ranges': all_data_ranges,
                'footer_config': footer_config,
                'all_tables_data': all_tables_data,
                'table_keys': table_keys,
                'mapping_rules': gt_layout_config.get('sheet_config', {}).get('data_flow', {}).get('mappings', {}),
                'DAF_mode': self.args.DAF,
                'override_total_text': None,
                'leather_summary': global_leather_summary
            }
        )
        
        return footer_builder.build()

    def _restore_template_footer(self, template_state_builder, current_row, table_keys):
        """
        Template footer restoration.
        Runs AFTER all tables and the Grand Total row are complete to place 
        static blueprint elements (signatures, warning text) beneath everything.
        """
        if template_state_builder and not self.sheet_config.get('skip_template_footer', False):
            try:
                # We need actual_num_cols which we can get from the last header_info
                actual_num_cols = getattr(self, 'header_info', {}).get('num_columns', None)
                
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
                    mode=gen_mode
                )
                logger.info("Template footer restored successfully")
            except Exception as e:
                logger.error(f"Failed to restore template footer: {e}", exc_info=True)
        else:
            logger.debug("Skipping template footer restoration (missing builder or config skip)")
