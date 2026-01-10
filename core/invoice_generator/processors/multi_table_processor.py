# invoice_generator/processors/multi_table_processor.py
import sys
import logging
import traceback
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter

from .base_processor import SheetProcessor
from ..builders.layout_builder import LayoutBuilder
from ..builders.footer_builder import FooterBuilder
from ..styling.models import StylingConfigModel, FooterData
from ..config.builder_config_resolver import BuilderConfigResolver
from ..builders.template_state_builder import TemplateStateBuilder
from ..utils.text_replacement_rules import build_replacement_rules
from ..extractors.header_extractor import HeaderExtractor

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

        # 2. Capture Template State
        template_state_builder = self._capture_template_state()
        if not template_state_builder:
            return False

        # 3. Initialize Tracking Variables
        structure_config = self.sheet_config.get('structure', {}) if self.sheet_config else {}
        current_row = structure_config.get('header_row', 21)
        all_data_ranges = []
        grand_total_pallets = 0
        last_header_info = None
        dynamic_desc_used = False
        
        # Use defaultdict for safer aggregation
        aggregated_leather_summary = defaultdict(lambda: defaultdict(float))

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
            (next_row, table_pallets, data_range, header_info, 
             table_dynamic_desc, table_leather_summary) = result
            
            # Update tracking
            current_row = next_row
            grand_total_pallets += table_pallets
            if data_range:
                all_data_ranges.append(data_range)
            last_header_info = header_info
            if table_dynamic_desc:
                dynamic_desc_used = True
            
            # Aggregate leather summary
            if table_leather_summary:
                for l_type, data in table_leather_summary.items():
                    for col_id, val in data.items():
                        aggregated_leather_summary[l_type][col_id] += val

        # 5. Build Grand Total Row
        if len(table_keys) > 1 and last_header_info:
            current_row = self._build_grand_total_row(
                current_row=current_row,
                grand_total_pallets=grand_total_pallets,
                all_data_ranges=all_data_ranges,
                last_header_info=last_header_info,
                all_tables_data=all_tables_data,
                table_keys=table_keys,
                aggregated_leather_summary=aggregated_leather_summary,
                dynamic_desc_used=dynamic_desc_used
            )

        # 6. Restore Template Footer
        self._restore_template_footer(template_state_builder, current_row, table_keys)
        
        logger.info(f"Successfully processed {len(table_keys)} tables for sheet '{self.sheet_name}'.")
        return True

    def _resolve_all_tables_data(self) -> Tuple[Optional[Dict], List]:
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
        if not all_tables_data or not isinstance(all_tables_data, dict):
            logger.warning(f"'processed_tables_data' not found/valid. Skipping '{self.sheet_name}'")
            return None, []

        table_keys = sorted(all_tables_data.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
        logger.info(f"Found {len(table_keys)} tables to process: {table_keys}")
        return all_tables_data, table_keys

    def _capture_template_state(self):
        """Captures template state (header/footer) for reuse."""
        from ..builders.template_state_builder import TemplateStateBuilder
        from ..builders.json_template_builder import JsonTemplateStateBuilder
        
        logger.info(f"[MultiTableProcessor] Capturing template state for reuse")
        
        # Check for JSON config first
        json_config = self.config_loader.get_template_json_config()
        if json_config and self.sheet_name in json_config:
            logger.info(f"Using JSON-based template state")
            try:
                sheet_layout_json = json_config[self.sheet_name]
                template_state_builder = JsonTemplateStateBuilder(
                    sheet_layout_data=sheet_layout_json,
                    debug=getattr(self.args, 'debug', False)
                )
                
                # Apply text replacements
                if self.args and self.invoice_data:
                    replacement_rules = build_replacement_rules(self.args)
                    template_state_builder.apply_text_replacements(
                        replacement_rules=replacement_rules,
                        invoice_data=self.invoice_data
                    )
                    self.replacements_log = template_state_builder.replacements_log
                    self.header_info = HeaderExtractor.extract(template_state_builder.header_state)
                    
                return template_state_builder
            except Exception as e:
                logger.critical(f"CRITICAL: JsonTemplateStateBuilder failed: {e}", exc_info=True)
                return None

        # Fallback to Legacy Excel Scan
        layout_config = self.sheet_config.get('layout_config', {}) if self.sheet_config else {}
        structure_config = layout_config.get('structure', {})
        
        if 'header_row' not in structure_config:
            logger.critical(f"CRITICAL: 'header_row' not found in sheet_config['layout_config']['structure'] for '{self.sheet_name}'. Cannot capture template state.")
            return None
        
        table_header_row = structure_config['header_row']
        template_header_end_row = table_header_row - 1
        template_footer_start_row = structure_config.get('footer_row', table_header_row + 1)
        num_header_cols = 20  # Conservative estimate
        
        try:
            template_state_builder = TemplateStateBuilder(
                worksheet=self.template_worksheet,
                num_header_cols=num_header_cols,
                header_end_row=template_header_end_row,
                footer_start_row=template_footer_start_row,
                debug=getattr(self.args, 'debug', False)
            )
            
            # Apply text replacements
            if self.args and self.invoice_data:
                try:
                    replacement_rules = build_replacement_rules(self.args)
                    template_state_builder.apply_text_replacements(
                        replacement_rules=replacement_rules,
                        invoice_data=self.invoice_data
                    )
                    # Capture replacements log
                    self.replacements_log = template_state_builder.replacements_log
                    
                    # Extract Header Info
                    self.header_info = HeaderExtractor.extract(template_state_builder.header_state)
                    
                except Exception as e:
                    logger.error(f"Failed to apply text replacements or extract header: {e}")
            
            return template_state_builder
        except Exception as e:
            logger.critical(f"CRITICAL: Failed to capture template state: {e}")
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
            enable_text_replacement=False, 
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
        
        layout_config['enable_text_replacement'] = False
        layout_config['skip_template_header_restoration'] = (not is_first_table)
        layout_config['skip_template_footer_restoration'] = True
        
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
            resolved_data.get('dynamic_desc_used', False),
            getattr(layout_builder, 'leather_summary', None)
        )

    def _build_grand_total_row(self, current_row, grand_total_pallets, all_data_ranges, last_header_info, 
                             all_tables_data, table_keys, aggregated_leather_summary, dynamic_desc_used):
        """Builds the Grand Total row after all tables."""
        logger.info("Adding Grand Total Row")
        
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
            leather_summary=dict(aggregated_leather_summary),
            weight_summary={'net': 0.0, 'gross': 0.0}  # Will be auto-filled with global weights by resolver
        )
        
        footer_builder = FooterBuilder(
            worksheet=self.output_worksheet,
            footer_data=footer_data,
            style_config={'styling_config': styling_model},
            context_config={
                'header_info': last_header_info,
                'pallet_count': grand_total_pallets,
                'sheet_name': self.sheet_name,
                'is_last_table': True,
                'dynamic_desc_used': dynamic_desc_used
            },
            data_config={
                'sum_ranges': all_data_ranges,
                'footer_config': footer_config,
                'all_tables_data': all_tables_data,
                'table_keys': table_keys,
                'mapping_rules': gt_layout_config.get('sheet_config', {}).get('data_flow', {}).get('mappings', {}),
                'DAF_mode': self.args.DAF,
                'override_total_text': None,
                'leather_summary': dict(aggregated_leather_summary)
            }
        )
        
        return footer_builder.build()

    def _restore_template_footer(self, template_state_builder, current_row, table_keys):
        """Restores the template footer at the end."""
        logger.info(f"[MultiTableProcessor] Restoring template footer after row {current_row}")
        
        actual_num_cols = None
        if table_keys:
            first_resolver = BuilderConfigResolver(
                config_loader=self.config_loader,
                sheet_name=self.sheet_name,
                worksheet=self.output_worksheet,
                args=self.args,
                invoice_data=self.invoice_data
            )
            _, _, first_layout_cfg = first_resolver.get_layout_bundles_with_data(table_key=table_keys[0])
            if first_layout_cfg and 'sheet_config' in first_layout_cfg:
                bundled_columns = first_layout_cfg['sheet_config'].get('structure', {}).get('columns', [])
                if bundled_columns:
                    actual_num_cols = len(bundled_columns)
        
        try:
            template_state_builder.restore_footer_only(
                target_worksheet=self.output_worksheet,
                footer_start_row=current_row,
                actual_num_cols=actual_num_cols
            )
        except Exception as e:
            logger.error(f"❌ Failed to restore template footer: {e}")
            logger.error(traceback.format_exc())
