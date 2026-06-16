import logging
import traceback
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, Optional, Union
from core.invoice_generator.models.footer import FooterData, WeightDetail, LeatherDetail
from ..bundle_accessor import BundleAccessor
from .base import TableSectionBuilder
from .grid import Grid

logger = logging.getLogger(__name__)


class TableFooterBuilder(BundleAccessor, TableSectionBuilder):
    """
    Builds and styles footer sections using pure bundle architecture via Grid.
    """
    
    def __init__(
        self,
        grid: Grid,
        footer_data: FooterData,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        data_config: Dict[str, Any]
    ):
        BundleAccessor.__init__(
            self,
            worksheet=None,  # Dropped raw worksheet
            style_config=style_config,
            context_config=context_config,
            data_config=data_config
        )
        TableSectionBuilder.__init__(self, grid)
        
        self.footer_data = footer_data
        self.initial_row = 0  # Grid is relative, start at 0 internally

    @property
    def sum_ranges(self) -> list:
        # Dynamically query data section from grid
        if hasattr(self.grid, "_sections") and "data" in self.grid._sections:
            start, end = self.grid.get_section_range("data")
            if start > 0 and end >= start:
                return [(start, end)]
        return self.data_config.get('sum_ranges', [])
    
    @property
    def footer_config(self) -> Dict[str, Any]:
        return self.data_config.get('footer_config', {})
    
    @property
    def pallet_count(self) -> int:
        return self.context_config.get('pallet_count', 0)
    
    @property
    def show_grand_total_addons(self) -> bool:
        return self.context_config.get('show_grand_total_addons', False)

    def build(self) -> None:
        logger.info(f"[FooterBuilder] build() called")
        if not self.footer_config:
            raise ValueError("[FooterBuilder] CANNOT BUILD FOOTER - Invalid config")

        try:
            current_footer_row = 0
            
            add_blank_before = self.footer_config.get("add_blank_before", False)
            if add_blank_before:
                current_footer_row += 1
            
            footer_type = self.footer_config.get("type", "regular")
            add_ons = self.footer_config.get("add_ons", {})
            
            before_footer_addon = add_ons.get("before_footer", {})
            if before_footer_addon.get("enabled", False) and footer_type == "regular":
                self._build_before_footer(current_footer_row, before_footer_addon)
                current_footer_row += 1
            
            self._build_main_footer(current_footer_row, footer_type)
            current_footer_row += 1

            if add_ons:
                current_footer_row = self._process_footer_addons(current_footer_row, add_ons, footer_type)

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
        # We handle grand_total by applying 'footer_no_border' context or we assume grid handles context well.
        context = 'footer' if footer_type != 'grand_total' else 'footer'
        
        written_cols = []

        # Write default cells
        footer_cells = self.footer_config.get("footer_cells", [])
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
        sum_column_ids = self.footer_config.get("sum_cols", [])
        if self.sum_ranges:
            for col_id in sum_column_ids:
                # We need to construct the formula inputs correctly for the grid.
                # Grid.write_formula expects template="=SUM({col_ref_0})" and inputs=["col_id"]
                # But sum_ranges are absolute rows like [(10, 15)].
                # Actually, our grid doesn't yet support absolute ranges in write_formula elegantly.
                # Let's just build the string. Wait, we don't know the column letter!
                try:
                    col_letter = self.grid.get_column_letter(col_id)
                    sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in self.sum_ranges]
                    formula = f"=SUM({','.join(sum_parts)})"
                    self.grid.write(row, col_id, formula, context=context)
                    written_cols.append(col_id)
                except ValueError:
                    pass

        self._pad_row_styles(row, exclude_cols=written_cols)

        merge_rules = self.footer_config.get("merge_rules", [])
        for rule in merge_rules:
            start_column_id = rule.get("start_column_id")
            colspan = rule.get("colspan")
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
        sum_column_ids = self.footer_config.get("sum_cols", [])

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
                # Get or create ensures the cell exists in the grid for export
                self.grid.write(row, col_id, None, context='footer')
