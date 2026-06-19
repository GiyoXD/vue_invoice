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

    def __init__(self, layout_config: Dict[str, Any]):
        """
        Initialize from layout configuration.

        Args:
            layout_config: Layout configuration containing:
                - structure: {row_heights: {context: float, ...}}
        """
        self._row_heights: Dict[str, Optional[float]] = {}

        structure_config = layout_config.get('structure', {}) if isinstance(layout_config, dict) else {}
        row_heights = structure_config.get('row_heights', {}) if isinstance(structure_config, dict) else {}
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
