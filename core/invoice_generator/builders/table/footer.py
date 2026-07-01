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
        logger.info(f"[FooterBuilder] build() called (declarative refactored engine)")
        if self.footer_config is None or not self.footer_config.rows:
            raise ValueError("[FooterBuilder] CANNOT BUILD FOOTER - Invalid config")

        try:
            rows_schema = self.footer_config.rows
            current_footer_row = 0

            # Prepare dynamic variables context payload
            pallet_count_multiple = "S" if self.pallet_count > 1 else ""
            
            weight_net = 0.0
            weight_gross = 0.0
            if self.footer_data and self.footer_data.weight_summary:
                ws = self.footer_data.weight_summary
                ws_dict = ws.model_dump() if hasattr(ws, 'model_dump') else ws
                weight_net = float(ws_dict.get('net', 0))
                weight_gross = float(ws_dict.get('gross', 0))
                
            payload = {
                "pallet_count": self.pallet_count,
                "multiple": pallet_count_multiple,
                "weight_net": weight_net,
                "weight_gross": weight_gross
            }
            
            # Populate leather summary variables
            if self.footer_data and self.footer_data.leather_summary:
                for l_type, summary_data in self.footer_data.leather_summary.items():
                    if hasattr(summary_data, 'model_dump'):
                        summary_data_dict = {
                            **summary_data.model_dump(by_alias=True),
                            **summary_data.model_dump(by_alias=False),
                            **(summary_data.model_extra or {})
                        }
                    else:
                        summary_data_dict = summary_data
                    
                    l_key = l_type.lower()
                    l_pallet = int(summary_data_dict.get('col_pallet_count', summary_data_dict.get('pallet_count', 0)))
                    payload[f"{l_key}_pallet_count"] = l_pallet
                    
                    for k, v in summary_data_dict.items():
                        payload[f"{l_key}_{k}"] = v

            # Keep track of active section boundaries
            active_section = None
            section_start_row = -1

            for row_schema in rows_schema:
                # Filter/Skip checks for specific dynamic addons
                is_row_skipped = False
                for cell in row_schema:
                    if cell.get("addon_type") == "leather":
                        leather_key = cell["leather_key"]
                        l_key = leather_key.lower()
                        pallet_val = payload.get(f"{l_key}_pallet_count", 0)
                        
                        has_sum = False
                        # Check if any sum keys in the payload are non-zero
                        for k, v in payload.items():
                            if k.startswith(f"{l_key}_") and k != f"{l_key}_pallet_count":
                                if v is not None and v != 0:
                                    has_sum = True
                                    break
                        
                        if pallet_val == 0 and not has_sum:
                            is_row_skipped = True
                            break
                            
                if is_row_skipped:
                    continue

                # Construct context-aware row payload
                row_payload = copy.deepcopy(payload)
                leather_cells = [c for c in row_schema if c.get("addon_type") == "leather"]
                if leather_cells:
                    l_key = leather_cells[0]["leather_key"].lower()
                    l_pallet = payload.get(f"{l_key}_pallet_count", 0)
                    row_payload["pallet_count"] = l_pallet
                    row_payload["multiple"] = "S" if l_pallet != 1 else ""
                    
                    # Copy all summary properties matching the prefix to their base names
                    prefix = f"{l_key}_"
                    for k, v in payload.items():
                        if k.startswith(prefix):
                            base_name = k[len(prefix):]
                            row_payload[base_name] = v

                # Determine style context from first cell
                row_context = row_schema[0].get("style_context", "footer") if row_schema else "footer"

                # If context rotates, close previous section and start new section
                if active_section != row_context:
                    if active_section is not None:
                        section_name = "grand_total" if self.footer_config and self.footer_config.type == "grand_total" and active_section == "footer" else active_section
                        self.grid.set_section_bounds(
                            section_name, 
                            self.grid._cursor_row + section_start_row, 
                            self.grid._cursor_row + current_footer_row - 1
                        )
                    active_section = row_context
                    section_start_row = current_footer_row

                written_cols = []
                
                # Render cells in the row
                for cell in row_schema:
                    col_id = cell["col_id"]
                    style_context = cell.get("style_context", "footer")
                    
                    # 1. Resolve formulas
                    if "formula" in cell:
                        func = cell["formula"]
                        target = cell.get("target_section", "data")
                        self.grid.write_section_aggregate(current_footer_row, col_id, function=func, section=target, context=style_context)
                        written_cols.append(col_id)
                        
                    # 1.5 Auto-lookup for leather addon without explicit value
                    elif cell.get("addon_type") == "leather" and "value" not in cell:
                        val = row_payload.get(col_id)
                        if val is not None:
                            if isinstance(val, str):
                                try:
                                    if '.' in val:
                                        val = float(val)
                                    else:
                                        val = int(val)
                                except ValueError:
                                    pass
                            self.grid.write(current_footer_row, col_id, val, context=style_context)
                            written_cols.append(col_id)
                        
                    # 2. Resolve normal text/numeric values
                    elif "value" in cell:
                        val = cell["value"]
                        if val is not None:
                            val_str = str(val)
                            try:
                                if "{" in val_str:
                                    val_str = val_str.format(**row_payload)
                            except Exception as format_err:
                                logger.debug(f"Placeholder formatting failed for text '{val_str}': {format_err}")
                                
                            if "{pallet_count}" in str(cell.get("value", "")) and self.pallet_count <= 0:
                                # Skip writing value for empty pallet count, but allow styling/merges
                                pass
                            elif cell.get("is_pallet") and val_str.isdigit():
                                self.grid.write(current_footer_row, col_id, int(val_str), context=style_context)
                                written_cols.append(col_id)
                            else:
                                try:
                                    if '.' in val_str:
                                        num_val = float(val_str)
                                    else:
                                        num_val = int(val_str)
                                    self.grid.write(current_footer_row, col_id, num_val, context=style_context)
                                except ValueError:
                                    self.grid.write(current_footer_row, col_id, val_str, context=style_context)
                                written_cols.append(col_id)

                    # 3. Apply merges
                    colspan = cell.get("colspan", 1)
                    if colspan > 1:
                        self.grid.merge(current_footer_row, col_id, rowspan=1, colspan=colspan)

                # Autopad missing columns to trigger styling
                self._pad_row_styles(current_footer_row, exclude_cols=written_cols, context=row_context)
                current_footer_row += 1

            # Close the last active section bounds
            if active_section is not None:
                section_name = "grand_total" if self.footer_config and self.footer_config.type == "grand_total" and active_section == "footer" else active_section
                self.grid.set_section_bounds(
                    section_name, 
                    self.grid._cursor_row + section_start_row, 
                    self.grid._cursor_row + current_footer_row - 1
                )

            self.grid.advance_row(current_footer_row)
            logger.info(f"[FooterBuilder] COMPLETE - generated {current_footer_row} rows")

        except Exception as e:
            logger.error(f"[FooterBuilder] FATAL ERROR during footer generation: {e}")
            logger.error(traceback.format_exc())
            raise

    def _pad_row_styles(self, row: int, exclude_cols=None, context: str = 'footer'):
        """Writes None to empty cells to trigger their background/border styles."""
        exclude_cols = exclude_cols or []
        excluded_idxs = set()
        for c_id in exclude_cols:
            idx = self.grid._resolve_column(c_id)
            if idx:
                excluded_idxs.add(idx)

        for col_id in self.grid.column_mapping.keys():
            col_idx = self.grid._resolve_column(col_id)
            if col_idx not in excluded_idxs:
                self.grid.write(row, col_id, None, context=context)
