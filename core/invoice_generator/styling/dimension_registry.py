"""
Dimension Registry - Row Height Lookup by Context

Provides row height values based on row context (header, data, footer).
Column widths are handled separately by LayoutBuilder reading from styling_config.

Pattern:
    registry = DimensionRegistry(sheet_config)
    height = registry.get_row_height('data')   # → 18.0
    height = registry.get_row_height('header') # → 30.0
"""

import logging
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


class DimensionRegistry:
    """
    Row height lookup by context (header, data, footer).

    This is intentionally minimal. Column widths are a sheet-level property
    handled by LayoutBuilder, not by the grid or dimension registry.

    Usage:
        registry = DimensionRegistry(sheet_config)
        height = registry.get_row_height('data')
    """

    def __init__(self, row_heights: Dict[str, float]):
        """
        Initialize from row heights dictionary.

        Args:
            row_heights: Row heights mapping dictionary.
        """
        self._row_heights: Dict[str, Optional[float]] = {}
        if isinstance(row_heights, dict):
            for context, height in row_heights.items():
                self._row_heights[context] = height

        logger.debug(
            f"DimensionRegistry loaded {len(self._row_heights)} contexts: "
            f"{list(self._row_heights.keys())}"
        )

    def get_row_height(self, context: str) -> Optional[float]:
        """Get row height for a specific context."""
        return self._row_heights.get(context)

    def has_context(self, context: str) -> bool:
        """Check if a context exists in the registry."""
        return context in self._row_heights
