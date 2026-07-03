import logging
from typing import Any, Dict, List, Optional, Union
import traceback

from core.invoice_generator.models.table_adapter import ResolvedTableData
from core.invoice_generator.config.table_value_adapter.preparer import _to_numeric
from .table_grid import Grid
from .base import TableSectionBuilder

logger = logging.getLogger(__name__)


class DataTableBuilderStyler(TableSectionBuilder):
    """
    Builds and styles data table sections based on pre-resolved data.
    
    This class is a "dumb" builder. Its only job is to take prepared data
    and write it to the grid. It does not contain any data-sourcing
    or mapping logic.
    """
    
    def __init__(
        self,
        grid: Grid,
        resolved_data: Union[ResolvedTableData, Dict[str, Any]],
        vertical_merge_columns: Optional[List[str]] = None,
        is_global_unique_desc: bool = False,
        parent_column_ids: Optional[List[str]] = None
    ):
        TableSectionBuilder.__init__(self, grid)
        self.parent_column_ids = parent_column_ids or []
        if isinstance(resolved_data, dict):
            # Pre-convert any integer keys in data_rows to strings to satisfy Pydantic key validation
            raw_rows = resolved_data.get('data_rows', [])
            if isinstance(raw_rows, list):
                stringified_rows = []
                for row in raw_rows:
                    if isinstance(row, dict):
                        stringified_rows.append({str(k): v for k, v in row.items()})
                    else:
                        stringified_rows.append(row)
                resolved_data['data_rows'] = stringified_rows
            resolved_data = ResolvedTableData.model_validate(resolved_data)
        self.resolved_data = resolved_data
        self.vertical_merge_columns = vertical_merge_columns or []
        self.is_global_unique_desc = is_global_unique_desc
 
        self.col_id_map = grid.column_mapping
        self.column_colspan = grid.column_colspan

        # Build a mapping from physical index to logical IDs to support integer keys in test cases
        idx_to_ids = {}
        for col_id, col_idx in self.col_id_map.items():
            if col_idx is not None:
                idx_to_ids.setdefault(col_idx, []).append(col_id)
                idx_to_ids.setdefault(str(col_idx), []).append(col_id)

        self.data_rows = []
        for row in resolved_data.data_rows:
            new_row = {}
            for k, v in row.items():
                if k in idx_to_ids:
                    for col_id in idx_to_ids[k]:
                        new_row[col_id] = v
                else:
                    new_row[k] = v
            self.data_rows.append(new_row)
        
        logger.debug(f"DataTableBuilder initialized with {len(self.data_rows)} total rows")

    def build(self) -> None:
        self.grid.mark_section_start("data")

        actual_rows_to_process = len(self.data_rows)
        num_data_rows = (self.resolved_data.num_data_rows or actual_rows_to_process) if hasattr(self.resolved_data, 'num_data_rows') else actual_rows_to_process
        
        try:
            for i in range(actual_rows_to_process):
                row_data = self.data_rows[i]
                
                # Filter row_data to only include columns in the valid col_id_map, excluding parent columns
                valid_col_ids = set(self.col_id_map.keys()) - set(self.parent_column_ids)
                row_data = {col_id: value for col_id, value in row_data.items() if col_id in valid_col_ids}
                
                columns_with_data = set(row_data.keys())

                # Write all columns for this row
                for col_id, value in row_data.items():
                    if isinstance(value, dict) and value.get('type') == 'formula':
                        self.grid.write_formula(i, col_id, value.get('template', ''), value.get('inputs', []), context='data')
                    else:
                        coerced_value = _to_numeric(value)
                        self.grid.write(i, col_id, coerced_value, context='data')
                
                # Handle columns defined in header but missing from row_data
                missing_columns = valid_col_ids - columns_with_data
                
                for col_id in missing_columns:
                    val = None
                    if col_id == 'col_no':
                        val = i + 1
                    
                    self.grid.write(i, col_id, val, context='data')

            # --- Apply Horizontal Merges ---
            if self.column_colspan:
                for r in range(actual_rows_to_process):
                    for col_id, colspan in self.column_colspan.items():
                        if colspan > 1:
                            self.grid.merge(r, col_id, rowspan=1, colspan=colspan)

            # --- Apply Vertical Merges ---
            if self.vertical_merge_columns and num_data_rows > 0:
                relative_start_row = 0
                relative_end_row = num_data_rows - 1
                
                for col_id in self.vertical_merge_columns:
                    if col_id not in self.col_id_map:
                        continue

                    if col_id == 'col_desc':
                        if not self.is_global_unique_desc:
                            logger.info("  Skipping vertical merge for col_desc because descriptions are mixed globally.")
                            continue
                    
                    # Validate desc baseline uniformity if col_id is col_desc
                    if col_id == "col_desc":
                        buffer_value = None
                        for r in range(relative_start_row, relative_end_row + 1):
                            cell = self.grid.get_cell(r, self.col_id_map[col_id])
                            if cell.value is not None and cell.value != "":
                                buffer_value = cell.value
                                break
                                
                        if buffer_value is not None:
                            baseline = str(buffer_value).strip().lower()
                            abort_merge = False
                            for r in range(relative_start_row, relative_end_row + 1):
                                cell = self.grid.get_cell(r, self.col_id_map[col_id])
                                if cell.value is not None and cell.value != "":
                                    current = str(cell.value).strip().lower()
                                    if current != baseline:
                                        abort_merge = True
                                        break
                            if abort_merge:
                                logger.info("  Aborting vertical merge for col_desc because values are mixed in this range.")
                                continue

                    group_start = relative_start_row
                    col_idx = self.col_id_map[col_id]
                    start_cell = self.grid.get_cell(relative_start_row, col_idx)
                    group_value = start_cell.value if start_cell else None
                    
                    for r in range(relative_start_row + 1, relative_end_row + 2):
                        if r <= relative_end_row:
                            cell = self.grid.get_cell(r, col_idx)
                            current_value = cell.value
                        else:
                            current_value = None  # sentinel to flush
                            
                        if current_value == group_value and r <= relative_end_row:
                            continue
                        else:
                            group_end = r - 1
                            if group_end > group_start and group_value is not None:
                                skip_int = False
                                try:
                                    int(group_value)
                                    skip_int = True
                                except (ValueError, TypeError):
                                    pass

                                if not skip_int:
                                    self.grid.merge(group_start, col_id, rowspan=group_end - group_start + 1, colspan=1)
                            
                            group_start = r
                            group_value = current_value

        except Exception as fill_data_err:
            logger.error(f"Error during data filling loop: {fill_data_err}\n{traceback.format_exc()}")
            raise

        # Advance grid row pointer
        self.grid.advance_row(actual_rows_to_process)
        self.grid.mark_section_end("data")
        logger.info(f"DataTableBuilder completed: {actual_rows_to_process} data rows generated")
