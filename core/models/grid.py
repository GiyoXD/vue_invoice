from typing import Any, Dict, List, Optional, Tuple
from core.models.cell import UnitRow, UnitCell, TemplateMerge, CellStyle

class Grid:
    """
    Core layout manager. Maps a sparse 2D matrix of physical rows and columns.
    Completely decoupled from logical business templates, style registries, or openpyxl.
    """
    def __init__(self):
        self._grid: Dict[int, Dict[int, UnitCell]] = {}
        self._row_heights: Dict[int, Optional[float]] = {}
        self._merge_map: Dict[Tuple[int, int], Tuple[int, int]] = {}

    def _get_or_create_cell(self, row: int, col: int) -> UnitCell:
        if row not in self._grid:
            self._grid[row] = {}
        if col not in self._grid[row]:
            self._grid[row][col] = UnitCell(col_index=col, value=None)
        return self._grid[row][col]

    def write(self, row: int, col: int, value: Any, style: Optional[CellStyle] = None, merge: Optional[TemplateMerge] = None):
        """
        Writes a value and optional styling/merge to the physical grid coordinates.
        """
        cell = self._get_or_create_cell(row, col)
        cell.value = value
        if style:
            cell.style = style
        if merge:
            cell.merge = merge

    def merge(self, row: int, col: int, rowspan: int = 1, colspan: int = 1):
        """
        Merges cells physically from the top-left (row, col) coordinates.
        Sets up the merge maps for lookup resolution.
        """
        if rowspan <= 1 and colspan <= 1:
            return

        cell = self._get_or_create_cell(row, col)
        cell.merge = TemplateMerge(
            min_col=col,
            max_col=col + colspan - 1,
            row_span=rowspan,
            value=str(cell.value) if cell.value is not None else ""
        )

        # Map all cells in this range to the top-left coordinator
        for r in range(row, row + rowspan):
            for c in range(col, col + colspan):
                self._merge_map[(r, c)] = (row, col)
                if r == row and c == col:
                    continue
                # Clear content and merge pointers of child cells to avoid duplication
                child_cell = self._get_or_create_cell(r, c)
                child_cell.value = None
                child_cell.merge = None

    def get_cell(self, row: int, col: int, resolve_merge: bool = False) -> UnitCell:
        """
        Gets a cell at the physical coordinate.
        If resolve_merge is True and the coordinate is within a merge range,
        it resolves to the top-left parent cell.
        """
        if resolve_merge and (row, col) in self._merge_map:
            tl_row, tl_col = self._merge_map[(row, col)]
            return self._get_or_create_cell(tl_row, tl_col)
        return self._get_or_create_cell(row, col)

    def set_row_height(self, row: int, height: Optional[float]):
        """Sets the height for a specific physical row index."""
        self._row_heights[row] = height

    def get_row_height(self, row: int) -> Optional[float]:
        """Gets the height for a physical row index."""
        return self._row_heights.get(row)

    def get_row_models(self) -> List[UnitRow]:
        """
        Exports the internal grid matrix to a list of UnitRow objects.
        """
        models = []
        if not self._grid:
            return models

        max_row = max(self._grid.keys())
        for r in range(max_row + 1):
            cells = []
            if r in self._grid:
                cells = sorted(list(self._grid[r].values()), key=lambda c: c.col_index)
            
            models.append(UnitRow(
                relative_index=r,
                height=self._row_heights.get(r),
                cells=cells
            ))
        return models

    def to_dict(self) -> Dict[str, Any]:
        """
        Serializes the grid structure, values, formatting, heights, and merge maps to a JSON-compatible dictionary.
        """
        # Convert _merge_map tuples keys to string keys for JSON compliance
        serialized_merge_map = {f"{r},{c}": list(tl) for (r, c), tl in self._merge_map.items()}
        return {
            "row_heights": {str(k): v for k, v in self._row_heights.items()},
            "merge_map": serialized_merge_map,
            "grid": {
                str(r): {str(c): cell.to_dict() for c, cell in row.items()}
                for r, row in self._grid.items()
            }
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'Grid':
        """
        Deserializes a JSON-compatible dictionary into a Grid instance.
        """
        grid_obj = cls()
        
        # Row heights
        row_heights_raw = d.get("row_heights", {})
        grid_obj._row_heights = {int(k): v for k, v in row_heights_raw.items()}
        
        # Merge map
        merge_map_raw = d.get("merge_map", {})
        for k, v in merge_map_raw.items():
            r, c = map(int, k.split(','))
            grid_obj._merge_map[(r, c)] = (v[0], v[1])
            
        # Grid cells
        grid_raw = d.get("grid", {})
        for r_str, row_data in grid_raw.items():
            r = int(r_str)
            grid_obj._grid[r] = {}
            for c_str, cell_data in row_data.items():
                c = int(c_str)
                grid_obj._grid[r][c] = UnitCell.from_dict(cell_data)
                
        return grid_obj
