import logging
import traceback
from typing import Any, Dict, Optional, List

from core.invoice_generator.models.config.layout import FooterConfigModel
from .base import TableSectionBuilder
from .table_grid import Grid

logger = logging.getLogger(__name__)


class TableFooterBuilder(TableSectionBuilder):
    """
    Builds and styles footer sections using pure bundle architecture via Grid.
    """
    
    def __init__(
        self,
        grid: Grid,
        footer_config: Optional[FooterConfigModel] = None,
        payload: Optional[Dict[str, Any]] = None,
        sum_ranges: Optional[List[tuple]] = None,
    ):
        super().__init__(grid)
        self.footer_config = footer_config or FooterConfigModel()
        self.payload = dict(payload) if isinstance(payload, dict) else {}
        self._custom_sum_ranges = sum_ranges

        self.initial_row = 0

    @property
    def sum_ranges(self) -> list:
        if self._custom_sum_ranges is not None:
            return self._custom_sum_ranges
        if hasattr(self.grid, "_sections") and "data" in self.grid._sections:
            start, end = self.grid.get_section_range("data")
            if start > 0 and end >= start:
                return [(start, end)]
        return []



    def build(self) -> None:
        logger.info(f"[FooterBuilder] build() called (declarative refactored engine)")
        if self.footer_config is None or not self.footer_config.rows:
            raise ValueError("[FooterBuilder] CANNOT BUILD FOOTER - Invalid config")

        try:
            flat_payload = self.payload.get("grand_total") if isinstance(self.payload.get("grand_total"), dict) else self.payload
            rows_rendered = self._build_declarative_section(
                rows_schema=self.footer_config.rows,
                payload=flat_payload,
                default_context="footer"
            )



            self.grid.advance_row(rows_rendered)
            logger.info(f"[FooterBuilder] COMPLETE - generated {rows_rendered} rows")

        except Exception as e:
            logger.error(f"[FooterBuilder] FATAL ERROR during footer generation: {e}")
            logger.error(traceback.format_exc())
            raise
