import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

from ...styling.style_registry import StyleRegistry
from ...models.layout import SheetLayoutState
from ...data.table_calculator import TableCalculator
from .grid import Grid

logger = logging.getLogger(__name__)


class TableBuilder:
    """
    Orchestrates building a single table grid (Header + Data + Footer).
    Delegates generation to sub-builders inheriting from TableSectionBuilder.
    """
    def __init__(
        self,
        workbook: Workbook,
        worksheet: Worksheet,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Dict[str, Any],
        layout_state: Optional[SheetLayoutState] = None
    ):
        self.workbook = workbook
        self.worksheet = worksheet
        self.style_config = style_config or {}
        self.context_config = context_config or {}
        self.layout_config = layout_config or {}
        self.layout_state = layout_state or SheetLayoutState()

        # Unpack styling and configuration
        self.styling_config = self.style_config.get('styling_config')
        self.sheet_config = self.layout_config.get('sheet_config', {})
        self.args = self.context_config.get('args')
        self.sheet_name = self.context_config.get('sheet_name', '')
        
        # Flags
        self.skip_header_builder = self.layout_config.get('skip_header_builder', False)
        self.skip_data_table_builder = self.layout_config.get('skip_data_table_builder', False)
        self.skip_footer_builder = self.layout_config.get('skip_footer_builder', False)
        
        # Extracted totals
        self.final_grand_total_pallets = self.context_config.get('final_grand_total_pallets', 0)
        self.total_net_weight = self.context_config.get('total_net_weight')
        self.total_gross_weight = self.context_config.get('total_gross_weight')
        self.is_last_table = self.context_config.get('is_last_table', False)
        self.show_grand_total_addons = self.context_config.get('show_grand_total_addons', False)

        # Output properties filled after building
        self.header_info = {}
        self.footer_data = None
        self.data_start_row = -1
        self.data_end_row = -1
        self.next_row_after_footer = -1

    def build(self, start_row: int) -> bool:
        """
        Builds the table at the given start_row index.
        Writes generated rows to the worksheet and returns success status.
        """
        # Local imports to break circular dependencies
        from .header import HeaderBuilderStyler as HeaderBuilder
        from .data import DataTableBuilderStyler as DataTableBuilder
        from .footer import TableFooterBuilder

        logger.info(f"[TableBuilder] Building table at row {start_row} for sheet '{self.sheet_name}'")
        
        # 1. Resolve columns (DAF/custom filters & index mappings)
        bundled_columns, column_mapping = self._resolve_columns()
        
        # 2. Setup StyleRegistry
        styling_dict = self.styling_config.model_dump() if hasattr(self.styling_config, 'model_dump') else self.styling_config
        style_registry = None
        if isinstance(styling_dict, dict) and 'columns' in styling_dict and 'row_contexts' in styling_dict:
            style_registry = StyleRegistry(styling_dict)

        # 3. Bind Layout State
        self.layout_state.bind(
            worksheet=self.worksheet,
            column_mapping=column_mapping,
            style_registry=style_registry
        )

        # 4. Initialize the Master Grid
        grid = Grid(column_mapping=column_mapping, style_registry=style_registry)
        grid.set_start_row(start_row)

        # 5. Build Table Header
        if not self.skip_header_builder and bundled_columns:
            try:
                header_builder = HeaderBuilder(
                    grid=grid,
                    start_row=start_row,
                    bundled_columns=bundled_columns
                )
                self.header_info = header_builder.build()
            except Exception as e:
                logger.error(f"[TableBuilder] HeaderBuilder crashed: {e}", exc_info=True)
                return False
        else:
            self.header_info = self.layout_config.get('header_info', {})
            if not self.header_info:
                self.header_info = {
                    'column_map': {},
                    'first_row_index': 0,
                    'second_row_index': 1,
                    'column_id_map': column_mapping
                }
            grid.advance_row(2)

        # 6. Build Data Table
        resolved_data = self.layout_config.get('resolved_data')
        data_physical_start_row = start_row + grid._cursor_row
        
        if not self.skip_data_table_builder and resolved_data:
            try:
                # Calculate metrics (TableCalculator)
                table_calculator = TableCalculator(self.header_info)
                self.footer_data = table_calculator.calculate(resolved_data)
                
                if not self.footer_data:
                    logger.error("[TableBuilder] TableCalculator failed")
                    return False
                
                # Calculate absolute boundaries for layout state recording
                actual_rows_to_process = len(resolved_data.get('data_rows', []))
                self.data_start_row = data_physical_start_row
                self.data_end_row = data_physical_start_row + actual_rows_to_process - 1
                
                if self.data_start_row > 0 and self.data_end_row >= self.data_start_row:
                    self.layout_state.record_data_range(self.data_start_row, self.data_end_row)

                # Uniqueness and vertical merge columns
                is_global_unique_desc = self.layout_config.get('is_global_unique_desc', False)
                allow_col_desc_merge = self.layout_config.get('allow_col_desc_merge', True)
                
                merge_cols = ['col_pallet_count']
                if allow_col_desc_merge:
                    merge_cols.append('col_desc')

                data_builder = DataTableBuilder(
                    grid=grid,
                    header_info=self.header_info,
                    resolved_data=resolved_data,
                    vertical_merge_columns=merge_cols,
                    is_global_unique_desc=is_global_unique_desc
                )
                data_builder.build()
                
            except Exception as e:
                logger.error(f"[TableBuilder] DataTableBuilder crashed: {e}", exc_info=True)
                return False
        else:
            self.data_start_row = 0
            self.data_end_row = 0

        # 7. Build Footer
        if not self.skip_footer_builder:
            pallet_count = self.footer_data.total_pallets if self.footer_data else self.final_grand_total_pallets
            footer_config = self.sheet_config.get('footer', {})
            data_flow = self.sheet_config.get('data_flow', {})
            mapping_rules = data_flow.get('mappings', self.sheet_config.get('mappings', {}))
            
            data_range_to_sum = []
            data_physical_end_row = start_row + grid._cursor_row - 1
            if data_physical_start_row > 0 and data_physical_end_row >= data_physical_start_row:
                data_range_to_sum = [(data_physical_start_row, data_physical_end_row)]

            footer_builder_context_config = {
                'header_info': self.header_info,
                'pallet_count': pallet_count,
                'sheet_name': self.sheet_name,
                'total_net_weight': self.total_net_weight,
                'total_gross_weight': self.total_gross_weight,
                'is_last_table': self.is_last_table,
                'show_grand_total_addons': self.show_grand_total_addons,
            }
            
            footer_builder_data_config = {
                'sum_ranges': data_range_to_sum,
                'footer_config': footer_config,
                'mapping_rules': mapping_rules,
                'DAF_mode': bool(getattr(self.args, 'DAF', False)) if self.args else False,
                'override_total_text': None,
                'leather_summary': self.footer_data.leather_summary if self.footer_data else None
            }

            try:
                footer_builder = TableFooterBuilder(
                    grid=grid,
                    footer_data=self.footer_data,
                    style_config={'styling_config': self.styling_config},
                    context_config=footer_builder_context_config,
                    data_config=footer_builder_data_config
                )
                footer_builder.build()
            except Exception as e:
                logger.error(f"[TableBuilder] TableFooterBuilder crashed: {e}", exc_info=True)
                return False

        # 8. Export and write to Layout State
        row_models = grid.get_row_models()
        if row_models:
            self.layout_state.write_row_models(row_models, start_row=start_row)

        self.next_row_after_footer = start_row + grid._cursor_row
        return True

    def _resolve_columns(self) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
        """
        Resolves filtered columns and builds the logical ID to physical column index mapping.
        """
        structure = self.sheet_config.get('structure', {})
        original_columns = structure.get('columns', [])
        
        column_mapping = {}
        bundled_columns = original_columns
        
        if original_columns:
            DAF_mode = self.args.DAF if self.args and hasattr(self.args, 'DAF') else False
            custom_mode = self.args.custom if self.args and hasattr(self.args, 'custom') else False
            
            template_col = 1
            output_col = 1
            
            for col_def in original_columns:
                skip_daf = bool(col_def.get('skip_in_daf', False))
                skip_custom = bool(col_def.get('skip_in_custom', False))
                colspan_val = int(col_def.get('colspan', 1))
                children_list = col_def.get('children', [])
                
                num_columns = len(children_list) if children_list else colspan_val
                should_skip = (DAF_mode and skip_daf) or (custom_mode and skip_custom)
                
                if should_skip:
                    for i in range(num_columns):
                        column_mapping[template_col + i] = None
                else:
                    for i in range(num_columns):
                        column_mapping[template_col + i] = output_col + i
                    output_col += num_columns
                
                template_col += num_columns

            # Filter columns list
            bundled_columns = [
                col for col in original_columns
                if not (DAF_mode and col.get('skip_in_daf', False))
                and not (custom_mode and col.get('skip_in_custom', False))
            ]

        # Convert logical ID to physical column mapping using filtered column layout
        resolved_col_id_map = {}
        if bundled_columns:
            col_index = 1
            for col in bundled_columns:
                col_id = col.get('id', '')
                if 'children' in col:
                    for child in col['children']:
                        child_id = child.get('id', '')
                        resolved_col_id_map[child_id] = col_index
                        col_index += 1
                else:
                    colspan = int(col.get('colspan', 1))
                    resolved_col_id_map[col_id] = col_index
                    col_index += colspan

        return bundled_columns, resolved_col_id_map
