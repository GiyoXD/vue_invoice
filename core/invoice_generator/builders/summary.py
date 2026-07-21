import logging
import copy
from typing import Any, Dict, List, Optional

from .table.base import TableSectionBuilder
from .table.table_grid import Grid
from core.invoice_generator.models.config.layout import FooterConfigModel

logger = logging.getLogger(__name__)


class SummaryBuilder(TableSectionBuilder):
    """
    Builds and styles page-level summary sections using pure bundle architecture via Grid.
    """
    def __init__(
        self,
        grid: Grid,
        summary_config: Optional[FooterConfigModel] = None,
        payload: Optional[Dict[str, Any]] = None,
        **kwargs
    ):
        super().__init__(grid)
        self.summary_config = summary_config or FooterConfigModel()
        self.payload = payload or {}

    def build(self) -> int:
        logger.info("[SummaryBuilder] build() called (declarative engine)")
        if self.summary_config is None or not self.summary_config.rows:
            logger.warning("[SummaryBuilder] CANNOT BUILD SUMMARY - Invalid config or no rows")
            return self.grid.start_row_index + self.grid._cursor_row

        try:
            rows_rendered = self._build_declarative_section(
                rows_schema=self.summary_config.rows,
                payload=self.payload,
                default_context="summary"
            )

            self.grid.advance_row(rows_rendered)
            logger.info(f"[SummaryBuilder] COMPLETE - generated {rows_rendered} rows in grid")
            return self.grid.start_row_index + self.grid._cursor_row

        except Exception as e:
            logger.exception("[SummaryBuilder] FATAL ERROR during summary generation")
            raise
