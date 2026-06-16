import logging
from typing import Any, Dict, List, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from ...utils.cell_converter import convert_registry_style_to_cell_style
from core.models.cell import UnitRow, UnitCell, TemplateMerge

from ...styling.models import StylingConfigModel
from ...styling.style_registry import StyleRegistry
from ...utils.layout import calculate_header_dimensions

from .grid import Grid
from .base import TableSectionBuilder

logger = logging.getLogger(__name__)


class HeaderBuilderStyler(TableSectionBuilder):
    def __init__(
        self,
        grid: Grid,
        start_row: int,
        bundled_columns: List[Dict[str, Any]]
    ):
        """
        Initialize HeaderBuilder with bundled config and a grid.
        
        Args:
            grid: The Grid instance to draw the header on.
            start_row: Starting row for header
            bundled_columns: Bundled format (list with id/header/format/rowspan/colspan/children)
        """
        self.start_row = start_row
        self.bundled_columns_original = bundled_columns
        
        TableSectionBuilder.__init__(self, grid)
        
        # Convert bundled columns to internal format
        if bundled_columns:
            logger.info(f"Using BUNDLED config (columns={len(bundled_columns)})")
            self.header_layout_config = self._convert_bundled_columns(bundled_columns)
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
        column_map = {}
        column_id_map = {}
        column_colspan = {}  # Track colspan for each column ID (excluding parents with children)
        
        # Identify parent columns (those with children) - they should NOT be in column_colspan
        parent_column_ids = set()
        if self.bundled_columns_original:
            for col in self.bundled_columns_original:
                if 'children' in col and col['children']:
                    parent_column_ids.add(col.get('id'))



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

                column_map[text] = get_column_letter(cell_col)
                column_id_map[cell_id] = cell_col
                # Only store colspan for NON-PARENT columns (parents with children shouldn't merge data/footer)
                if cell_id not in parent_column_ids:
                    column_colspan[cell_id] = colspan

        self.grid.advance_row(num_header_rows) if hasattr(self.grid, 'advance_row') else None
        self.grid.mark_section_end("header")
    
    def _convert_bundled_columns(self, columns: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Convert bundled columns format to internal header_layout_config format.
        
        Bundled format:
            {"id": "col_po", "header": "P.O. №", "format": "@", "rowspan": 2}
        
        Internal format:
            {"row": 0, "col": 1, "text": "P.O. №", "id": "col_po", "rowspan": 2, "colspan": 1}
        """
        headers = []
        col_index = 0
        
        for col in columns:
            col_id = col.get('id', '')
            header_text = col.get('header', '')
            rowspan = col.get('rowspan', 1)
            colspan = col.get('colspan', 1)
            
            # Handle parent column with children (e.g., Quantity with PCS/SF)
            if 'children' in col:
                # Add parent header
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': 1,
                    'colspan': len(col['children'])
                })
                
                # Add children headers
                for child in col['children']:
                    headers.append({
                        'row': 1,
                        'col': col_index,
                        'text': child.get('header', ''),
                        'id': child.get('id', ''),
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
