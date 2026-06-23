import logging
from typing import Any, Dict, List, Optional, Tuple

from core.invoice_generator.models.config.layout import ColumnDef
from ...utils.layout import calculate_header_dimensions

from .table_grid import Grid
from .base import TableSectionBuilder

logger = logging.getLogger(__name__)


class HeaderBuilderStyler(TableSectionBuilder):
    def __init__(
        self,
        grid: Grid,
        start_row: int,
        bundled_columns: List[ColumnDef]
    ):
        """
        Initialize HeaderBuilder with bundled config and a grid.
        
        Args:
            grid: The Grid instance to draw the header on.
            start_row: Starting row for header
            bundled_columns: List of ColumnDef objects
        """
        self.start_row = start_row
        
        # Parse dicts to ColumnDef if necessary
        parsed_columns = []
        if bundled_columns:
            for col in bundled_columns:
                if isinstance(col, dict):
                    parsed_columns.append(ColumnDef.model_validate(col))
                else:
                    parsed_columns.append(col)
                    
        self.bundled_columns_original = parsed_columns
        
        TableSectionBuilder.__init__(self, grid)
        
        # Convert bundled columns to internal format
        if parsed_columns:
            logger.info(f"Using BUNDLED config (columns={len(parsed_columns)})")
            self.header_layout_config = self._convert_bundled_columns(parsed_columns)
            logger.debug(f"Converted to {len(self.header_layout_config)} header cells")
        else:
            logger.error("HeaderBuilder: No bundled columns provided!")
            raise ValueError("No bundled columns provided")

    def build(self) -> None:
        if not self.header_layout_config or self.start_row <= 0:
            return None

        self.grid.mark_section_start("header")

        num_header_rows, num_header_cols = calculate_header_dimensions(self.header_layout_config)

        first_row_index = self.start_row
        last_row_index = self.start_row
        max_col = 0
        column_colspan = {}  # Track colspan for each column ID (excluding parents with children)
        
        # Identify parent columns (those with children) - they should NOT be in column_colspan
        parent_column_ids = set()
        if self.bundled_columns_original:
            for col in self.bundled_columns_original:
                if col.children:
                    parent_column_ids.add(col.id)

        for cell_config in self.header_layout_config:
            row_offset = cell_config.get('row', 0)
            col_offset = cell_config.get('col', 0)
            text = cell_config.get('text', '')
            cell_id = cell_config.get('id')
            rowspan = cell_config.get('rowspan', 1)
            colspan = cell_config.get('colspan', 1)

            cell_row = self.start_row + row_offset
            cell_col = 1 + col_offset

            last_row_index = max(last_row_index, cell_row + rowspan - 1)
            max_col = max(max_col, cell_col + colspan - 1)

            # Write to grid
            if cell_id:
                self.grid.write(row_offset, cell_id, text, context='header')
                if rowspan > 1 or colspan > 1:
                    self.grid.merge(row_offset, cell_id, rowspan, colspan)

                # Only store colspan for NON-PARENT columns (parents with children shouldn't merge data/footer)
                if cell_id not in parent_column_ids:
                    column_colspan[cell_id] = colspan

        self.grid.advance_row(num_header_rows) if hasattr(self.grid, 'advance_row') else None
        self.grid.mark_section_end("header")
    
    def _convert_bundled_columns(self, columns: List[ColumnDef]) -> List[Dict[str, Any]]:
        """
        Convert bundled columns format to internal header_layout_config format.
        """
        headers = []
        col_index = 0
        
        for col in columns:
            col_id = col.id
            header_text = col.header
            rowspan = col.rowspan
            colspan = col.colspan
            
            # Handle parent column with children (e.g., Quantity with PCS/SF)
            if col.children:
                # Add parent header
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': 1,
                    'colspan': len(col.children)
                })
                
                # Add children headers
                for child in col.children:
                    headers.append({
                        'row': 1,
                        'col': col_index,
                        'text': child.header,
                        'id': child.id,
                        'rowspan': 1,
                        'colspan': 1
                    })
                    col_index += 1
            else:
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': rowspan,
                    'colspan': colspan
                })
                # Increment by colspan to skip physical columns occupied by merge
                col_index += colspan
        
        return headers
