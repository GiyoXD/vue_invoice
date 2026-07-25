import logging
import copy
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional

from .table_grid import Grid

logger = logging.getLogger(__name__)


class TableSectionBuilder(ABC):
    """
    Abstract base class for all table section builders (Header, Data, Footer).
    Workers now operate purely on a Grid, abstracting away physical coordinates.
    """
    def __init__(self, grid: Grid):
        self.grid = grid

    @abstractmethod
    def build(self) -> Any:
        """Build the section using the grid."""
        pass

    def _render_row(self, row_schema: List[Dict[str, Any]], row_payload: Dict[str, Any], current_row: int, style_context: str) -> List[str]:
        written_cols = []
        for cell in row_schema:
            col_id = cell["col_id"]
            cell_style = cell.get("style_context", style_context)
            
            # 1. Resolve formulas
            if "formula" in cell:
                func = cell["formula"]
                target = cell.get("target_section", "data")
                self.grid.write_section_aggregate(
                    current_row, 
                    col_id, 
                    function=func, 
                    section=target, 
                    context=cell_style,
                    sum_ranges=getattr(self, 'sum_ranges', [])
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
                    cell_val_template = str(cell.get("value", ""))
                    if ("{pallet_count}" in cell_val_template or "{col_pallet_count}" in cell_val_template) and (row_payload.get("col_pallet_count", 0) <= 0):
                        pass
                    elif cell.get("is_pallet") and val_str.isdigit():
                        self.grid.write(current_row, col_id, int(val_str), context=cell_style)
                        written_cols.append(col_id)
                    else:
                        try:
                            if '.' in val_str:
                                num_val = float(val_str)
                            else:
                                num_val = int(val_str)
                            self.grid.write(current_row, col_id, num_val, context=cell_style)
                        except ValueError:
                            self.grid.write(current_row, col_id, val_str, context=cell_style)
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
                    self.grid.write(current_row, col_id, val, context=cell_style)
                    written_cols.append(col_id)
            
            # 4. Apply merges
            colspan = cell.get("colspan", 1)
            if colspan > 1:
                self.grid.merge(current_row, col_id, rowspan=1, colspan=colspan)
                
        return written_cols

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

    def _build_declarative_section(
        self,
        rows_schema: List[Any],
        payload: Dict[str, Any],
        default_context: str = "footer",
        section_name_override: Optional[str] = None
    ) -> int:
        """
        Shared declarative row rendering engine.
        Handles repeating rows (dict with source_list) and standard single rows (list of cells).
        Returns number of rows rendered.
        """
        current_row = 0
        active_section = None
        section_start_row = -1
        cursor_base = self.grid._cursor_row

        for row_item in rows_schema:
            if isinstance(row_item, dict):
                if "source_list" not in row_item:
                    logger.error(f"Invalid row_item dict structure (missing source_list): {row_item}")
                    continue

                source_list_key = row_item["source_list"]
                cells_schema = row_item.get("cells", [])
                records = payload.get(source_list_key) or []

                for record in records:
                    fallback_context = default_context if default_context == "summary" else f"{default_context}_addon"
                    row_context = cells_schema[0].get("style_context", fallback_context) if cells_schema else fallback_context

                    if active_section != row_context:
                        if active_section is not None:
                            final_name = section_name_override if section_name_override and active_section == default_context else active_section
                            self.grid.set_section_bounds(
                                final_name,
                                cursor_base + section_start_row,
                                cursor_base + current_row - 1
                            )
                        active_section = row_context
                        section_start_row = current_row

                    rec_payload = self._enrich_payload(record)
                    written_cols = self._render_row(cells_schema, rec_payload, current_row, row_context)
                    self._pad_row_styles(current_row, exclude_cols=written_cols, context=row_context)
                    current_row += 1

            elif isinstance(row_item, list):
                row_schema = row_item
                if not row_schema:
                    continue

                row_context = row_schema[0].get("style_context", default_context) if row_schema else default_context

                if active_section != row_context:
                    if active_section is not None:
                        final_name = section_name_override if section_name_override and active_section == default_context else active_section
                        self.grid.set_section_bounds(
                            final_name,
                            cursor_base + section_start_row,
                            cursor_base + current_row - 1
                        )
                    active_section = row_context
                    section_start_row = current_row

                row_payload = self._prepare_row_payload(payload, row_schema)
                written_cols = self._render_row(row_schema, row_payload, current_row, row_context)
                self._pad_row_styles(current_row, exclude_cols=written_cols, context=row_context)
                current_row += 1
            else:
                logger.error(f"Invalid row_item structure (expected dict or list): {row_item}")
                continue

        if active_section is not None:
            final_name = section_name_override if section_name_override and active_section == default_context else active_section
            self.grid.set_section_bounds(
                final_name,
                cursor_base + section_start_row,
                cursor_base + current_row - 1
            )

        return current_row

    def _enrich_payload(self, record: Dict[str, Any]) -> Dict[str, Any]:
        """Standard payload enrichment for repeating row records."""
        rec_payload = copy.deepcopy(record)
        pallet_count = rec_payload.pop("pallet_count", rec_payload.get("col_pallet_count", 0))
        rec_payload["col_pallet_count"] = pallet_count
        rec_payload["multiple"] = "S" if pallet_count != 1 else ""
        rec_payload["weight_net"] = rec_payload.get("col_net", rec_payload.get("net", 0.0))
        rec_payload["weight_gross"] = rec_payload.get("col_gross", rec_payload.get("gross", 0.0))
        return rec_payload

    def _prepare_row_payload(self, base_payload: Dict[str, Any], row_schema: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Prepare payload for standard rows, handling legacy leather addon fields and prefix remapping."""
        row_payload = copy.deepcopy(base_payload)
        p_count = row_payload.pop("pallet_count", row_payload.get("col_pallet_count", 0))
        row_payload["col_pallet_count"] = p_count
        row_payload["pallet_count"] = p_count
        row_payload.setdefault("multiple", "S" if p_count != 1 else "")

        leather_cells = [c for c in row_schema if isinstance(c, dict) and c.get("addon_type") == "leather"]
        if leather_cells:
            l_key = leather_cells[0]["leather_key"].lower()
            l_pallet = base_payload.get(f"{l_key}_pallet_count", 0)
            row_payload["pallet_count"] = l_pallet
            row_payload["multiple"] = "S" if l_pallet != 1 else ""

            prefix = f"{l_key}_"
            for k, v in base_payload.items():
                if k.startswith(prefix):
                    base_name = k[len(prefix):]
                    row_payload[base_name] = v

        return row_payload

