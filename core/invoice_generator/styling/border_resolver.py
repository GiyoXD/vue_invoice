"""
Border Resolver - Centralized Border Logic for Table Grids

Stamps borders onto a completed TableGrid as a post-build step.
All border decision-making lives here — no other module makes border choices.

Usage:
    resolver = BorderResolver(
        default_border="side_grid",
        column_overrides={"col_static": "sides_only"}
    )
    # After all sections are built:
    resolver.apply(grid)

Resolution (per cell):
    1. Column override (e.g., col_static → sides_only)
    2. Default border mode (side_grid → sides_only, full_grid → thin)
    3. Section context (header/footer always get full border)
    4. Position: last row of grid always gets bottom border to close table
"""

import logging
from copy import copy
from typing import Dict, Optional

from core.models.cell import BorderStyle, CellStyle

logger = logging.getLogger(__name__)


# All patterns use "thin" — per template spec. Tables never use medium/thick.
BORDER_PATTERNS: Dict[str, Optional[BorderStyle]] = {
    "thin": BorderStyle(left="thin", right="thin", top="thin", bottom="thin"),
    "sides_only": BorderStyle(left="thin", right="thin", top=None, bottom=None),
    "no_bottom": BorderStyle(left="thin", right="thin", top="thin", bottom=None),
    "no_top": BorderStyle(left="thin", right="thin", top=None, bottom="thin"),
    "none": None,
}

# Contexts that always get full thin borders regardless of default_border mode
FULL_BORDER_CONTEXTS = {"header", "footer"}
NO_BORDER_CONTEXTS = {"footer_addon", "grand_total"}


class BorderResolver:
    """
    Stamps borders onto a completed TableGrid.

    Called ONCE after all sections (header, data, footer) are built.
    Walks the grid range and applies borders based on:
      - Column config (per-column pattern overrides like sides_only)
      - Default border mode (side_grid / full_grid)
      - Section context (header/footer always full border)
      - Position (last row of grid always gets a bottom border)
    """

    def __init__(self, default_border: str = "full_grid", column_overrides: Optional[Dict[str, str]] = None):
        """
        Args:
            default_border: "full_grid" (all borders) or "side_grid" (sides only for data)
            column_overrides: Per-column border pattern names.
                              e.g., {"col_static": "sides_only", "col_price": "thin"}
        """
        self.default_border = default_border
        self.column_overrides = column_overrides or {}

    def _get_base_pattern(self, col_id: str, context: str) -> Optional[BorderStyle]:
        """
        Determines the base border pattern for a cell before positional adjustments.

        Priority:
            1. NO_BORDER_CONTEXTS (always None)
            2. Header context -> always full thin (no overrides)
            3. Column override (explicit per-column pattern in data/footer rows)
            4. Footer/before-footer context -> always full thin
            5. Default border mode (side_grid → sides_only, full_grid → thin)
        """
        if context in NO_BORDER_CONTEXTS:
            return copy(BORDER_PATTERNS["none"])

        # Header context ALWAYS gets full thin borders, bypassing column overrides
        if context == "header":
            return copy(BORDER_PATTERNS["thin"])

        # Column-level override takes precedence in other contexts (data, footer)
        if col_id in self.column_overrides:
            pattern_name = self.column_overrides[col_id]
            # Normalize legacy naming
            if pattern_name == "side_only":
                pattern_name = "sides_only"
            pattern = BORDER_PATTERNS.get(pattern_name)
            return copy(pattern) if pattern else None

        # Footer / before-footer always get full borders
        if context in FULL_BORDER_CONTEXTS:
            return copy(BORDER_PATTERNS["thin"])

        # Default border mode for data rows
        if self.default_border == "side_grid":
            return copy(BORDER_PATTERNS["sides_only"])
        else:
            # full_grid or any other value → full thin
            return copy(BORDER_PATTERNS["thin"])

    def _get_context_for_row(self, grid, row: int) -> str:
        """
        Determines which section context a row belongs to.
        Uses the grid's section markers (header, data, footer).
        Both grid keys and section ranges are relative (0-based).
        """
        for section_name, (start, end) in grid._sections.items():
            if start <= row <= end:
                return section_name
        # Default to data if no section matches
        return "data"

    def apply(self, grid) -> None:
        """
        Post-build: stamps borders onto every cell in the completed grid.

        Walks all populated rows, determines the appropriate border pattern
        for each cell, and ensures the last row always has a bottom border.
        """
        if not grid._grid:
            return

        all_rows = sorted(grid._grid.keys())
        
        # Find the absolute last row of the "actual" table (excluding borderless add-ons)
        last_table_row = -1
        for section, (start, end) in grid._sections.items():
            if section not in NO_BORDER_CONTEXTS:
                last_table_row = max(last_table_row, end)
        
        # Fallback if somehow no standard sections are present
        if last_table_row == -1:
            last_table_row = all_rows[-1]

        for row in all_rows:
            context = self._get_context_for_row(grid, row)

            for col_id, col_idx in grid.column_mapping.items():
                if row not in grid._grid or col_idx not in grid._grid[row]:
                    continue

                cell = grid._grid[row][col_idx]
                pattern = self._get_base_pattern(col_id, context)

                # Position override: last table row always gets a bottom border to close the grid
                if row == last_table_row and pattern is not None and pattern.bottom is None:
                    pattern = BorderStyle(
                        left=pattern.left,
                        right=pattern.right,
                        top=pattern.top,
                        bottom="thin"
                    )

                # Stamp border onto cell
                if pattern is not None:
                    if cell.style is None:
                        cell.style = CellStyle(border=pattern)
                    else:
                        cell.style.border = pattern

        logger.debug(
            f"BorderResolver applied borders to {len(all_rows)} rows "
            f"(mode={self.default_border}, overrides={list(self.column_overrides.keys())})"
        )
