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

    def _render_row(self, row_schema: List[Dict[str, Any]], row_payload: Dict[str, Any], current_footer_row: int, style_context: str) -> List[str]:
        written_cols = []
        for cell in row_schema:
            col_id = cell["col_id"]
            cell_style = cell.get("style_context", style_context)
            
            # 1. Resolve formulas
            if "formula" in cell:
                func = cell["formula"]
                target = cell.get("target_section", "data")
                self.grid.write_section_aggregate(
                    current_footer_row, 
                    col_id, 
                    function=func, 
                    section=target, 
                    context=cell_style,
                    sum_ranges=self.sum_ranges
                )
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
                        
                    # Skip writing empty pallet count description
                    if "{pallet_count}" in str(cell.get("value", "")) and row_payload.get("pallet_count", 0) <= 0:
                        pass
                    elif cell.get("is_pallet") and val_str.isdigit():
                        self.grid.write(current_footer_row, col_id, int(val_str), context=cell_style)
                        written_cols.append(col_id)
                    else:
                        try:
                            if '.' in val_str:
                                num_val = float(val_str)
                            else:
                                num_val = int(val_str)
                            self.grid.write(current_footer_row, col_id, num_val, context=cell_style)
                        except ValueError:
                            self.grid.write(current_footer_row, col_id, val_str, context=cell_style)
                        written_cols.append(col_id)
            
            # 3. Direct key lookup from payload using col_id (col_id language)
            else:
                val = row_payload.get(col_id)
                if val is not None:
                    try:
                        if isinstance(val, str):
                            if '.' in val:
                                val = float(val)
                            elif val.isdigit():
                                val = int(val)
                    except ValueError:
                        pass
                    self.grid.write(current_footer_row, col_id, val, context=cell_style)
                    written_cols.append(col_id)
            
            # 4. Apply merges
            colspan = cell.get("colspan", 1)
            if colspan > 1:
                self.grid.merge(current_footer_row, col_id, rowspan=1, colspan=colspan)
                
        return written_cols

    def build(self) -> None:
        logger.info(f"[FooterBuilder] build() called (declarative refactored engine)")
        if self.footer_config is None or not self.footer_config.rows:
            raise ValueError("[FooterBuilder] CANNOT BUILD FOOTER - Invalid config")

        try:
            rows_schema = self.footer_config.rows
            current_footer_row = 0

            # Keep track of active section boundaries
            active_section = None
            section_start_row = -1

            for row_item in rows_schema:
                # Determine if this is a repeating row
                if isinstance(row_item, dict) and "source_list" in row_item:
                    source_list_key = row_item["source_list"]
                    cells_schema = row_item.get("cells", [])
                    records = self.payload.get(source_list_key, [])
                    
                    for record in records:
                        # Determine style context from first cell
                        row_context = cells_schema[0].get("style_context", "footer_addon") if cells_schema else "footer_addon"
                        
                        # Handle section change
                        if active_section != row_context:
                            if active_section is not None:
                                section_name = "grand_total" if self.footer_config.type == "grand_total" and active_section == "footer" else active_section
                                self.grid.set_section_bounds(
                                    section_name,
                                    self.grid._cursor_row + section_start_row,
                                    self.grid._cursor_row + current_footer_row - 1
                                )
                            active_section = row_context
                            section_start_row = current_footer_row
                            
                        # Format record variables (like multiple, pallet_count, weight_net, weight_gross if not present)
                        rec_payload = copy.deepcopy(record)
                        pallet_count = rec_payload.get("col_pallet_count", rec_payload.get("pallet_count", 0))
                        rec_payload["pallet_count"] = pallet_count
                        rec_payload["multiple"] = "S" if pallet_count != 1 else ""
                        rec_payload["weight_net"] = rec_payload.get("col_net", rec_payload.get("net", 0.0))
                        rec_payload["weight_gross"] = rec_payload.get("col_gross", rec_payload.get("gross", 0.0))
                        
                        written_cols = self._render_row(cells_schema, rec_payload, current_footer_row, row_context)
                        
                        # Autopad missing columns to trigger styling
                        self._pad_row_styles(current_footer_row, exclude_cols=written_cols, context=row_context)
                        current_footer_row += 1
                else:
                    # Standard single footer row
                    row_schema = row_item
                    if not row_schema:
                        continue
                        
                    # Determine style context from first cell
                    row_context = row_schema[0].get("style_context", "footer") if row_schema else "footer"
                    
                    # Handle section change
                    if active_section != row_context:
                        if active_section is not None:
                            section_name = "grand_total" if self.footer_config.type == "grand_total" and active_section == "footer" else active_section
                            self.grid.set_section_bounds(
                                section_name,
                                self.grid._cursor_row + section_start_row,
                                self.grid._cursor_row + current_footer_row - 1
                            )
                        active_section = row_context
                        section_start_row = current_footer_row
                        
                    # Support legacy leather addon fields in standard rows by copying prefixed keys to base keys
                    row_payload = copy.deepcopy(self.payload)
                    leather_cells = [c for c in row_schema if c.get("addon_type") == "leather"]
                    if leather_cells:
                        l_key = leather_cells[0]["leather_key"].lower()
                        l_pallet = self.payload.get(f"{l_key}_pallet_count", 0)
                        row_payload["pallet_count"] = l_pallet
                        row_payload["multiple"] = "S" if l_pallet != 1 else ""
                        
                        # Copy all summary properties matching the prefix to their base names
                        prefix = f"{l_key}_"
                        for k, v in self.payload.items():
                            if k.startswith(prefix):
                                base_name = k[len(prefix):]
                                row_payload[base_name] = v
                                
                    written_cols = self._render_row(row_schema, row_payload, current_footer_row, row_context)
                    
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
