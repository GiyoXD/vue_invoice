import logging
import traceback
import copy
from decimal import Decimal
from typing import Any, Dict, Optional, List

from core.invoice_generator.models.config.layout import FooterConfigModel
from core.invoice_generator.models.footer import FooterData
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
        footer_data: Optional[FooterData] = None,
        **kwargs
    ):
        TableSectionBuilder.__init__(self, grid)
        self.footer_config = footer_config or FooterConfigModel()
        self.payload = payload or {}
        
        # Check kwargs for footer_data if it wasn't passed positionally
        if footer_data is None:
            footer_data = kwargs.get('footer_data')
            
        # Backward compatibility conversion
        if footer_data and not payload:
            self.payload = {
                "pallet_count": footer_data.total_pallets,
                "multiple": "S" if footer_data.total_pallets != 1 else "",
            }
            if footer_data.weight_summary:
                ws = footer_data.weight_summary
                ws_dict = ws.model_dump() if hasattr(ws, 'model_dump') else ws
                self.payload["weight_net"] = float(ws_dict.get('net', 0.0))
                self.payload["weight_gross"] = float(ws_dict.get('gross', 0.0))
                
            if footer_data.leather_summary:
                leather_records = []
                for l_type, l_data in footer_data.leather_summary.items():
                    l_dict = l_data.model_dump() if hasattr(l_data, 'model_dump') else l_data
                    p_count = l_dict.get('col_pallet_count', l_dict.get('pallet_count', 0))
                    
                    l_key = l_type.lower()
                    # Populate old flat keys for legacy test compatibility
                    self.payload[f"{l_key}_pallet_count"] = p_count
                    for k, v in l_dict.items():
                        self.payload[f"{l_key}_{k}"] = v
                        if not k.startswith("col_"):
                            self.payload[f"{l_key}_col_{k}"] = v
                    
                    qty_sf = l_dict.get('col_qty_sf', l_dict.get('col_qty', 0.0))
                    if p_count == 0 and qty_sf == 0:
                        continue
                    record = {
                        "leather_type": l_type,
                        "pallet_count": int(p_count),
                        "col_qty_pcs": int(l_dict.get('col_qty_pcs', l_dict.get('col_qty', 0))),
                        "col_qty_sf": float(l_dict.get('col_qty_sf', l_dict.get('col_qty', 0.0))),
                        "col_net": float(l_dict.get('col_net', 0.0)),
                        "col_gross": float(l_dict.get('col_gross', 0.0)),
                        "col_cbm": float(l_dict.get('col_cbm', 0.0)),
                    }
                    leather_records.append(record)
                self.payload["leather_summary"] = leather_records
                
        self._pallet_count_override = kwargs.get('pallet_count')
        if self._pallet_count_override is not None:
            self.payload["pallet_count"] = self._pallet_count_override
            self.payload["multiple"] = "S" if self._pallet_count_override != 1 else ""
            
        self.initial_row = 0  # Grid is relative, start at 0 internally
        self._custom_sum_ranges = sum_ranges

    @property
    def pallet_count(self) -> int:
        if self._pallet_count_override is not None:
            return self._pallet_count_override
        return self.payload.get('pallet_count', 0)

    @property
    def sum_ranges(self) -> list:
        if self._custom_sum_ranges is not None:
            return self._custom_sum_ranges
        # Dynamically query data section from grid
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
            section_override = "grand_total" if self.footer_config.type == "grand_total" else None

            rows_rendered = self._build_declarative_section(
                rows_schema=self.footer_config.rows,
                payload=self.payload,
                default_context="footer",
                section_name_override=section_override
            )

            self.grid.advance_row(rows_rendered)
            logger.info(f"[FooterBuilder] COMPLETE - generated {rows_rendered} rows")

        except Exception as e:
            logger.error(f"[FooterBuilder] FATAL ERROR during footer generation: {e}")
            logger.error(traceback.format_exc())
            raise



