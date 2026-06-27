import logging
import traceback
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
        footer_data: Optional[FooterData] = None,
        footer_config: Optional[FooterConfigModel] = None,
        pallet_count: int = 0,
        show_grand_total_addons: bool = False,
        is_daf: bool = False,
        sheet_name: str = "",
        sum_ranges: Optional[List[tuple]] = None,
        **kwargs
    ):
        TableSectionBuilder.__init__(self, grid)
        self.footer_data = footer_data
        
        # Extract from legacy test dictionary parameters if provided
        data_config = kwargs.get('data_config', {}) or {}
        context_config = kwargs.get('context_config', {}) or {}
        
        if footer_config is None:
            raw_footer = data_config.get('footer_config', {})
            if isinstance(raw_footer, dict):
                footer_config = FooterConfigModel.model_validate(raw_footer)
            elif isinstance(raw_footer, FooterConfigModel):
                footer_config = raw_footer
            else:
                footer_config = FooterConfigModel()
                
        if not pallet_count and 'pallet_count' in context_config:
            pallet_count = context_config['pallet_count']
            
        if not sheet_name and 'sheet_name' in context_config:
            sheet_name = context_config['sheet_name']
            
        if not is_daf:
            if 'is_daf' in context_config:
                is_daf = context_config['is_daf']
            elif 'DAF_mode' in data_config:
                is_daf = data_config['DAF_mode']
                
        if sum_ranges is None and 'sum_ranges' in data_config:
            sum_ranges = data_config['sum_ranges']

        self.footer_config = footer_config
        self.pallet_count = pallet_count
        self.show_grand_total_addons = show_grand_total_addons
        self.is_daf = is_daf
        self.sheet_name = sheet_name
        self.initial_row = 0  # Grid is relative, start at 0 internally
        self._custom_sum_ranges = sum_ranges

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
        logger.info(f"[FooterBuilder] build() called")
        if self.footer_config is None or (
            self.footer_config.total_text_column_id is None and
            not self.footer_config.sum_cols and
            not self.footer_config.sum_column_ids and
            not self.footer_config.footer_cells and
            not self.footer_config.merge_rules
        ):
            raise ValueError("[FooterBuilder] CANNOT BUILD FOOTER - Invalid config")

        try:
            current_footer_row = 0
            
            add_blank_before = self.footer_config.add_blank_before
            if add_blank_before:
                current_footer_row += 1
            
            footer_type = self.footer_config.type or "regular"
            add_ons = self.footer_config.add_ons or {}
            
            before_footer_addon = add_ons.get("before_footer", {})
            if before_footer_addon.get("enabled", False) and footer_type == "regular":
                start_row = current_footer_row
                self._build_before_footer(current_footer_row, before_footer_addon)
                current_footer_row += 1
                self.grid.set_section_bounds("footer_addon_before", self.grid._cursor_row + start_row, self.grid._cursor_row + current_footer_row - 1)
            
            start_row = current_footer_row
            self._build_main_footer(current_footer_row, footer_type)
            current_footer_row += 1
            self.grid.set_section_bounds("footer", self.grid._cursor_row + start_row, self.grid._cursor_row + current_footer_row - 1)

            if add_ons:
                start_row = current_footer_row
                current_footer_row = self._process_footer_addons(current_footer_row, add_ons, footer_type)
                if current_footer_row > start_row:
                    self.grid.set_section_bounds("footer_addon", self.grid._cursor_row + start_row, self.grid._cursor_row + current_footer_row - 1)

            self.grid.advance_row(current_footer_row)
            logger.info(f"[FooterBuilder] COMPLETE - generated {current_footer_row} rows")

        except Exception as e:
            logger.error(f"[FooterBuilder] FATAL ERROR during footer generation: {e}")
            logger.error(traceback.format_exc())
            raise

    def _build_before_footer(self, row: int, config: Dict[str, Any]):
        column_id = config.get('column_id')
        text = config.get('text', '')
        merge_span = config.get('merge', 0)
        
        if not column_id or not text:
            return
        
        self.grid.write(row, column_id, text, context='footer')
        if merge_span > 1:
            self.grid.merge(row, column_id, rowspan=1, colspan=merge_span)
            
        # Write empty values to trigger styling for all other columns
        self._pad_row_styles(row, exclude_cols=[column_id])

    def _build_main_footer(self, row: int, footer_type: str):
        context = 'footer'
        written_cols = []

        # Write default cells
        footer_cells = self.footer_config.footer_cells or []
        for cell_config in footer_cells:
            if not isinstance(cell_config, list) or len(cell_config) < 2:
                continue
                
            text = str(cell_config[0])
            col_id = cell_config[1]
            
            if "{pallet_count}" in text:
                if self.pallet_count <= 0:
                    continue
                text = text.replace("{pallet_count}", str(self.pallet_count))
                text = text.replace("{multiple}", "S" if self.pallet_count > 1 else "")
                
            self.grid.write(row, col_id, text, context=context)
            written_cols.append(col_id)

        # Write sum formulas
        sum_column_ids = self.footer_config.sum_cols or []
        if self.sum_ranges:
            for col_id in sum_column_ids:
                try:
                    col_letter = self.grid.get_column_letter(col_id)
                    sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in self.sum_ranges]
                    formula = f"=SUM({','.join(sum_parts)})"
                    self.grid.write(row, col_id, formula, context=context)
                    written_cols.append(col_id)
                except ValueError:
                    pass

        self._pad_row_styles(row, exclude_cols=written_cols)

        merge_rules = self.footer_config.merge_rules or []
        for rule in merge_rules:
            start_column_id = rule.start_column_id
            colspan = rule.colspan
            if start_column_id and colspan:
                self.grid.merge(row, start_column_id, rowspan=1, colspan=colspan)

    def _process_footer_addons(self, start_row: int, add_ons: dict, footer_type: str) -> int:
        current_row = start_row
        
        weight_config = add_ons.get("weight_summary", {})
        if weight_config.get("enabled"):
            if self.show_grand_total_addons or footer_type == "grand_total":
                current_row = self._build_weight_summary_addon(current_row, weight_config)
        
        leather_config = add_ons.get("leather_summary", {})
        if leather_config.get("enabled"):
            if self.show_grand_total_addons or footer_type == "grand_total":
                current_row = self._build_leather_summary_addon(current_row, leather_config)
                
        return current_row

    def _build_weight_summary_addon(self, row: int, config: Dict[str, Any]) -> int:
        label_col_id = config.get("label_col_id")
        value_col_id = config.get("value_col_id")
        
        if not label_col_id or not value_col_id:
            return row
            
        grand_total_net = Decimal('0')
        grand_total_gross = Decimal('0')
        
        if self.footer_data and self.footer_data.weight_summary:
            ws = self.footer_data.weight_summary
            ws_dict = ws.model_dump() if hasattr(ws, 'model_dump') else ws
            try:
                grand_total_net = Decimal(str(ws_dict.get('net', 0)))
                grand_total_gross = Decimal(str(ws_dict.get('gross', 0)))
            except Exception:
                pass

        # NW Row
        self.grid.write(row, label_col_id, "NW(KGS)", context='footer')
        self.grid.write(row, value_col_id, float(grand_total_net), context='footer')
        self._pad_row_styles(row, exclude_cols=[label_col_id, value_col_id])
        
        # GW Row
        self.grid.write(row + 1, label_col_id, "GW(KGS):", context='footer')
        self.grid.write(row + 1, value_col_id, float(grand_total_gross), context='footer')
        self._pad_row_styles(row + 1, exclude_cols=[label_col_id, value_col_id])

        return row + 2

    def _build_leather_summary_addon(self, row: int, config: Dict[str, Any]) -> int:
        if self.sheet_name != "Packing list":
            return row
            
        leather_summary = self.footer_data.leather_summary if self.footer_data else None
        if not leather_summary:
            return row

        current_row = row
        sum_column_ids = self.footer_config.sum_cols or []

        # Get configurable column IDs with backward-compatible defaults
        label_col_id = config.get("label_col_id", "col_desc")
        pallet_col_id = config.get("pallet_col_id", "col_pallet_count")

        for leather_type in ['BUFFALO', 'COW']:
            summary_data = leather_summary.get(leather_type)
            if not summary_data:
                continue

            if hasattr(summary_data, 'model_dump'):
                summary_data_dict = {
                    **summary_data.model_dump(by_alias=True),
                    **summary_data.model_dump(by_alias=False),
                    **(summary_data.model_extra or {})
                }
            else:
                summary_data_dict = summary_data

            pallet_count = int(summary_data_dict.get('col_pallet_count', summary_data_dict.get('pallet_count', 0)))
            has_sum = any(col_id in summary_data_dict for col_id in sum_column_ids)

            if pallet_count == 0 and not has_sum:
                continue

            written_cols = []

            # Write label
            type_text = "LEATHER" if leather_type == 'COW' else f"{leather_type} LEATHER"
            self.grid.write(current_row, label_col_id, type_text, context='footer')
            written_cols.append(label_col_id)

            # Write values
            if pallet_count > 0:
                self.grid.write(current_row, pallet_col_id, str(pallet_count), context='footer')
                written_cols.append(pallet_col_id)

            for col_id in sum_column_ids:
                if col_id in summary_data_dict:
                    self.grid.write(current_row, col_id, summary_data_dict[col_id], context='footer')
                    written_cols.append(col_id)

            self._pad_row_styles(current_row, exclude_cols=written_cols)
            current_row += 1

        return current_row

    def _pad_row_styles(self, row: int, exclude_cols=None):
        """Writes None to empty cells to trigger their background/border styles."""
        exclude_cols = exclude_cols or []
        for col_id in self.grid.column_mapping.keys():
            if col_id not in exclude_cols:
                self.grid.write(row, col_id, None, context='footer')
