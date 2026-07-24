import logging
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

from ...styling.style_registry import StyleRegistry
from ...styling.dimension_registry import DimensionRegistry
from ...models.layout import SheetLayoutState
from ...models.config.styling import SheetStylingModel
from ...models.config.layout import SheetLayoutModel, ColumnDef, FooterConfigModel
from ...mappers import ResolvedTableData, resolve_summary_payload
from .table_grid import Grid


logger = logging.getLogger(__name__)


@dataclass
class TableBuilderConfig:
    worksheet: Worksheet
    sheet_styling: SheetStylingModel
    sheet_layout: SheetLayoutModel
    resolved_data: ResolvedTableData


class TableBuilder:
    """
    Orchestrates building a single table grid (Header + Data + Footer).
    Delegates generation to sub-builders inheriting from TableSectionBuilder.
    """
    def __init__(self, config: TableBuilderConfig):
        self.config = config

        # Output properties filled after building
        self.header_info = {}
        self.footer_data = None
        self.data_start_row = -1
        self.data_end_row = -1
        self.next_row_after_footer = -1

    def build(self, start_row: int, layout_state: SheetLayoutState) -> bool:
        """
        Builds the table at the given start_row index.
        Writes generated rows to the worksheet and returns success status.
        """
        # Local imports to break circular dependencies
        from .header import HeaderBuilderStyler as HeaderBuilder
        from .data import DataTableBuilderStyler as DataTableBuilder
        from .footer import TableFooterBuilder

        self.layout_state = layout_state

        logger.info(f"[TableBuilder] Building table at row {start_row}")
        
        # 1. Resolve columns (DAF/custom filters & index mappings)
        bundled_columns, column_mapping, column_colspan = self._resolve_columns()
        
        # 2. Setup StyleRegistry and DimensionRegistry
        style_registry = StyleRegistry(self.config.sheet_styling)
        row_heights = {
            context: style.row_height
            for context, style in self.config.sheet_styling.row_contexts.items()
            if style.row_height is not None
        }
        dimension_registry = DimensionRegistry(row_heights)

        # 3. Bind Layout State
        self.layout_state.bind(
            worksheet=self.config.worksheet,
            column_mapping=column_mapping,
            style_registry=style_registry,
            dimension_registry=dimension_registry
        )

        # 4. Initialize the Master Grid
        grid = Grid(column_mapping=column_mapping, style_registry=style_registry, column_colspan=column_colspan, dimension_registry=dimension_registry)
        grid.column_index_mapping = getattr(self, 'column_index_mapping', {})
        grid.set_start_row(start_row)
        self.grid = grid

        # 5. Build Table Header
        if bundled_columns:
            try:
                header_builder = HeaderBuilder(
                    grid=grid,
                    start_row=start_row,
                    bundled_columns=bundled_columns
                )
                header_builder.build()
            except Exception as e:
                logger.error(f"[TableBuilder] HeaderBuilder crashed: {e}", exc_info=True)
                return False
        else:
            grid.advance_row(2)

        # 6. Build Data Table
        data_physical_start_row = start_row + grid._cursor_row
        
        if self.config.resolved_data and self.config.resolved_data.data_rows:
            try:
                # Directly construct FooterData from pre-calculated parser results
                from ...models.footer import FooterData
                footer = self.config.resolved_data.footer
                pallet_count = footer.pallet_summary_total if footer else 0
                if pallet_count is None:
                    pallet_count = 0
                    
                ws = footer.weight_summary if footer else None

                self.footer_data = FooterData(
                    footer_row_start_idx=data_physical_start_row + len(self.config.resolved_data.data_rows),
                    data_start_row=data_physical_start_row,
                    data_end_row=data_physical_start_row + len(self.config.resolved_data.data_rows) - 1,
                    total_pallets=int(pallet_count),
                    leather_summary=footer.leather_summary if footer else None,
                    weight_summary=ws
                )
                
                # Calculate absolute boundaries for layout state recording
                actual_rows_to_process = len(self.config.resolved_data.data_rows)
                self.data_start_row = data_physical_start_row
                self.data_end_row = data_physical_start_row + actual_rows_to_process - 1
                
                if self.data_start_row > 0 and self.data_end_row >= self.data_start_row:
                    self.layout_state.record_data_range(self.data_start_row, self.data_end_row)

                parent_column_ids = [col.id for col in bundled_columns if col.children]

                data_builder = DataTableBuilder(
                    grid=grid,
                    resolved_data=self.config.resolved_data,
                    parent_column_ids=parent_column_ids
                )
                data_builder.build()
                
            except Exception as e:
                logger.error(f"[TableBuilder] DataTableBuilder crashed: {e}", exc_info=True)
                return False
        else:
            self.data_start_row = 0
            self.data_end_row = 0

        # 7. Build Footer
        if self.config.sheet_layout.footer:
            try:
                # 7a. Build Decoupled HS Code Row
                if self.config.sheet_layout.hs_code:
                    hs_model = self.config.sheet_layout.hs_code
                    grid.write(0, hs_model.col_id, hs_model.value, context=hs_model.style_context)
                    if hs_model.colspan > 1:
                        grid.merge(0, hs_model.col_id, rowspan=1, colspan=hs_model.colspan)
                    
                    # Pad other columns in this row to trigger styling
                    written_idxs = {grid._resolve_column(hs_model.col_id)}
                    for col_id in grid.column_mapping.keys():
                        c_idx = grid._resolve_column(col_id)
                        if c_idx not in written_idxs:
                            grid.write(0, col_id, None, context=hs_model.style_context)
                    
                    grid.advance_row(1)
            except Exception as e:
                logger.error(f"[TableBuilder] HS Code Builder crashed: {e}", exc_info=True)
                return False

            # Build the payload using resolve_summary_payload
            footer = self.config.resolved_data.footer if self.config.resolved_data else None
            invoice_data_for_payload = None
            if footer:
                invoice_data_for_payload = {
                    'footer_data': {
                        'grand_total': footer.grand_total or {},
                        'leather_summary': footer.leather_summary or []
                    }
                }

            payload = resolve_summary_payload(
                invoice_data=invoice_data_for_payload,
                footer_data_model=self.footer_data
            )


            try:
                footer_builder = TableFooterBuilder(
                    grid=grid,
                    footer_config=self.config.sheet_layout.footer or FooterConfigModel(),
                    payload=payload
                )
                footer_builder.build()
            except Exception as e:
                logger.error(f"[TableBuilder] TableFooterBuilder crashed: {e}", exc_info=True)
                return False

        # 8. Finalize borders (post-build: stamps borders onto the completed grid)
        from ...styling.border_resolver import BorderResolver
        
        # Collect column-level border overrides from config
        column_border_overrides = {}
        for col_id, col_def in self.config.sheet_styling.columns.items():
            if col_def.border_style:
                border_style = col_def.border_style
                # Normalize legacy naming
                if border_style == 'side_only':
                    border_style = 'sides_only'
                column_border_overrides[col_id] = border_style
        
        border_resolver = BorderResolver(
            default_border=self.config.sheet_styling.default_border or "full_grid",
            column_overrides=column_border_overrides
        )
        border_resolver.apply(grid)

        # 9. Export and write to Layout State
        row_models = grid.get_row_models()
        if row_models:
            self.layout_state.write_row_models(row_models, start_row=start_row)

        self.next_row_after_footer = start_row + grid._cursor_row
        return True

    def _resolve_columns(self) -> Tuple[List[ColumnDef], Dict[str, int], Dict[str, int]]:
        """
        Resolves filtered columns and builds the logical ID to physical column index mapping.
        """
        self.column_index_mapping = self.config.sheet_layout.column_index_mapping
        return self.config.sheet_layout.bundled_columns, self.config.sheet_layout.column_mapping, self.config.sheet_layout.column_colspan
