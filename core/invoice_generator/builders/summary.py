import logging
from typing import Any, Dict, List, Optional

from .table.base import TableSectionBuilder
from .table.table_grid import Grid
from core.invoice_generator.models.config.layout import FooterConfigModel
from core.invoice_generator.styling.border_resolver import apply_border_resolver

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
    ):
        super().__init__(grid)
        self.summary_config = summary_config or FooterConfigModel()
        self.payload = dict(payload) if isinstance(payload, dict) else {}

    def build(self) -> int:
        if self.summary_config is None or not self.summary_config.rows:
            logger.warning("[SummaryBuilder] Skipped build: summary_config is missing or has 0 rows.")
            return self.grid.start_row_index + self.grid._cursor_row

        try:
            flat_payload = dict(self.payload)
            ws = self.payload.get("weight_summary")
            if isinstance(ws, dict):
                flat_payload.update(ws)

            logger.info(f"[SummaryBuilder] Building summary with payload keys: {list(flat_payload.keys())}")
            rows_rendered = self._build_declarative_section(
                rows_schema=self.summary_config.rows,
                payload=flat_payload,
                default_context="summary"
            )



            if rows_rendered == 0:
                logger.warning("[SummaryBuilder] Built 0 rows. Check if source_list keys exist in payload.")
            else:
                logger.info(f"[SummaryBuilder] Successfully rendered {rows_rendered} summary rows.")

            # Stamp borders onto summary section grid cells
            apply_border_resolver(self.grid)

            self.grid.advance_row(rows_rendered)
            return self.grid.start_row_index + self.grid._cursor_row

        except Exception as e:
            logger.exception("[SummaryBuilder] FATAL ERROR during summary generation")
            raise

