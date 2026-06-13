import logging
from typing import Any, Dict, List, Optional, Tuple, Union
from openpyxl.worksheet.worksheet import Worksheet
from ..utils.cell_converter import convert_registry_style_to_cell_style
from core.models.cell import UnitRow, UnitCell, TemplateMerge, AlignmentStyle
from openpyxl.utils import get_column_letter
import traceback

logger = logging.getLogger(__name__)

from ..data.data_preparer import prepare_data_rows, parse_mapping_rules
from ..utils.layout import apply_column_widths, merge_contiguous_cells_by_id
from ..styling.style_registry import StyleRegistry
from ..styling.models import StylingConfigModel
from .bundle_accessor import BundleAccessor

class DataTableBuilderStyler:
    """
    Builds and styles data table sections based on pre-resolved data.
    
    This class is a "dumb" builder. Its only job is to take prepared data
    and write it to the worksheet. It does not contain any data-sourcing
    or mapping logic.
    """
    
    def __init__(
        self,
        worksheet: Worksheet,
        header_info: Dict[str, Any],
        resolved_data: Dict[str, Any],
        sheet_styling_config: Optional[StylingConfigModel] = None,
        vertical_merge_columns: Optional[List[str]] = None,
        is_global_unique_desc: bool = False
    ):
        """
        Initialize the builder with resolved data.
        
        Args:
            worksheet: The worksheet to write to.
            header_info: Header information with column maps.
            resolved_data: The data prepared by TableDataAdapter.
            sheet_styling_config: The styling configuration for the sheet.
            is_global_unique_desc: Whether descriptions are unique across the entire dataset.
        """
        self.worksheet = worksheet
        self.header_info = header_info
        self.resolved_data = resolved_data
        self.sheet_styling_config = sheet_styling_config
        self.vertical_merge_columns = vertical_merge_columns or []
        self.is_global_unique_desc = is_global_unique_desc

        # Extract commonly used values
        self.data_rows = resolved_data.get('data_rows', [])
        self.static_info = resolved_data.get('static_info', {})
        self.formula_rules = resolved_data.get('formula_rules', {})
        self.pallet_counts = resolved_data.get('pallet_counts', [])
        
        self.col_id_map = header_info.get('column_id_map', {})
        self.idx_to_id_map = {v: k for k, v in self.col_id_map.items()}
        self.column_colspan = header_info.get('column_colspan', {})  # Colspan for automatic merging
        
        # Initialize StyleRegistry for ID-driven styling
        self.style_registry = None
        if sheet_styling_config:
            try:
                styling_dict = sheet_styling_config.model_dump() if hasattr(sheet_styling_config, 'model_dump') else sheet_styling_config
                if isinstance(styling_dict, dict) and 'columns' in styling_dict and 'row_contexts' in styling_dict:
                    self.style_registry = StyleRegistry(styling_dict)
                    logger.info("StyleRegistry initialized successfully for DataTableBuilder")
                else:
                    logger.error(f"DataTableBuilder: Invalid styling config format. Expected 'columns' and 'row_contexts'.")
                    raise ValueError("Invalid styling config format")
            except Exception as e:
                logger.error(f"Could not initialize StyleRegistry: {e}")
                raise
        else:
            logger.error("DataTableBuilder: No styling config provided!")
            raise ValueError("No styling config provided")
        
        # Static content is now injected into data_rows by TableDataResolver
        # No need to handle it separately here
        logger.debug(f"DataTableBuilder initialized with {len(self.data_rows)} total rows (including any static rows)")

    def build(self) -> Optional[List[UnitRow]]:
        if not self.header_info or 'second_row_index' not in self.header_info:
            logger.error("Invalid header_info provided to DataTableBuilderStyler")
            return None

        num_columns = self.header_info.get('num_columns', 0)
        data_writing_start_row = self.header_info.get('second_row_index', 0) + 1
        
        actual_rows_to_process = len(self.data_rows)
        
        data_start_row = data_writing_start_row
        data_end_row = data_start_row + actual_rows_to_process - 1 if actual_rows_to_process > 0 else data_start_row - 1
        
        # Calculate actual end row of valid items, excluding static row padding
        num_data_rows = self.resolved_data.get('num_data_rows', actual_rows_to_process)
        actual_data_end_row = data_start_row + num_data_rows - 1 if num_data_rows > 0 else data_start_row - 1
        
        # --- Fill Data Rows Loop ---
        try:
            grid = {r: {} for r in range(actual_rows_to_process)}

            for i in range(actual_rows_to_process):
                current_row_idx = data_start_row + i
                row_data = self.data_rows[i]
                
                # Filter row_data to only include columns in the filtered column_id_map
                valid_col_indices = set(self.col_id_map.values())
                row_data = {col_idx: value for col_idx, value in row_data.items() if col_idx in valid_col_indices}
                
                columns_with_data = set(row_data.keys())

                # Write all columns for this row
                for col_idx, value in row_data.items():
                    val = None
                    if isinstance(value, dict) and value.get('type') == 'formula':
                        val = self._build_formula_string(value, current_row_idx)
                    else:
                        if isinstance(value, str):
                            if not value.strip():
                                val = None
                            else:
                                try:
                                    float_val = float(value)
                                    if float_val.is_integer():
                                        val = int(float_val)
                                    else:
                                        val = float_val
                                except (ValueError, TypeError):
                                    val = value
                        else:
                            val = value
                    
                    # Apply styling using StyleRegistry if available
                    col_id = self.idx_to_id_map.get(col_idx)
                    if not col_id:
                        logger.error(f"❌ Column index {col_idx} has NO column ID mapping!")
                        continue
                    
                    if not self.style_registry:
                        logger.error(f"❌ StyleRegistry not initialized!")
                        continue
                    
                    style_dict = self.style_registry.get_style(col_id, context='data')
                    
                    if col_id == 'col_static':
                        from copy import deepcopy
                        style_dict = deepcopy(style_dict)
                        style_dict['border_style'] = 'sides_only'
                    
                    cell_style = convert_registry_style_to_cell_style(style_dict)
                    grid[i][col_idx] = UnitCell(col_index=col_idx, value=val, style=cell_style)
                
                # Handle columns defined in header but missing from row_data
                all_column_indices = set(self.col_id_map.values())
                missing_columns = all_column_indices - columns_with_data
                
                for col_idx in missing_columns:
                    col_id = self.idx_to_id_map.get(col_idx)
                    val = None
                    
                    if col_id == 'col_no':
                        val = i + 1
                    
                    if not self.style_registry:
                        logger.error(f"❌ StyleRegistry not initialized for column {col_id}")
                        continue
                    
                    style_dict = self.style_registry.get_style(col_id, context='data')
                    
                    if col_id == 'col_static':
                        from copy import deepcopy
                        style_dict = deepcopy(style_dict)
                        style_dict['border_style'] = 'sides_only'
                    
                    cell_style = convert_registry_style_to_cell_style(style_dict)
                    grid[i][col_idx] = UnitCell(col_index=col_idx, value=val, style=cell_style)

            # --- Apply Horizontal Merges (based on colspan from header structure) ---
            if self.column_colspan:
                for r in range(actual_rows_to_process):
                    for col_id, colspan in self.column_colspan.items():
                        if colspan > 1:
                            col_idx = self.col_id_map.get(col_id)
                            if col_idx and col_idx in grid[r]:
                                start_col = col_idx
                                end_col = col_idx + colspan - 1
                                grid[r][col_idx].merge = TemplateMerge(
                                    min_col=start_col,
                                    max_col=end_col,
                                    row_span=1,
                                    value=str(grid[r][col_idx].value) if grid[r][col_idx].value is not None else ""
                                )
                                # Clear value of other cells in this colspan range, preserving style
                                for clear_c in range(start_col + 1, end_col + 1):
                                    if clear_c in grid[r]:
                                        grid[r][clear_c].value = None

            # --- Apply Vertical Merges (Conditional based on GLOBAL col_desc uniqueness) ---
            if self.vertical_merge_columns and num_data_rows > 0:
                relative_start_row = 0
                relative_end_row = num_data_rows - 1
                
                for col_id in self.vertical_merge_columns:
                    col_idx = self.col_id_map.get(col_id)
                    if not col_idx:
                        continue

                    if col_id == 'col_desc':
                        if self.is_global_unique_desc:
                            pass
                        else:
                            logger.info("  Skipping vertical merge for col_desc because descriptions are mixed globally.")
                            continue
                    
                    # Validate desc baseline uniformity if col_id is col_desc
                    if col_id == "col_desc":
                        buffer_value = None
                        for r in range(relative_start_row, relative_end_row + 1):
                            if r in grid and col_idx in grid[r]:
                                val = grid[r][col_idx].value
                                if val is not None and val != "":
                                    buffer_value = val
                                    break
                        if buffer_value is not None:
                            baseline = str(buffer_value).strip().lower()
                            abort_merge = False
                            for r in range(relative_start_row, relative_end_row + 1):
                                if r in grid and col_idx in grid[r]:
                                    val = grid[r][col_idx].value
                                    if val is not None and val != "":
                                        current = str(val).strip().lower()
                                        if current != baseline:
                                            abort_merge = True
                                            break
                            if abort_merge:
                                logger.info("  Aborting vertical merge for col_desc because values are mixed in this range.")
                                continue

                    group_start = relative_start_row
                    group_value = grid[relative_start_row][col_idx].value if (relative_start_row in grid and col_idx in grid[relative_start_row]) else None
                    
                    for r in range(relative_start_row + 1, relative_end_row + 2):
                        if r <= relative_end_row:
                            current_value = grid[r][col_idx].value if (r in grid and col_idx in grid[r]) else None
                        else:
                            current_value = None  # Sentinel to flush
                            
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
                                    if group_start in grid and col_idx in grid[group_start]:
                                        grid[group_start][col_idx].merge = TemplateMerge(
                                            min_col=col_idx,
                                            max_col=col_idx,
                                            row_span=group_end - group_start + 1,
                                            value=str(group_value)
                                        )
                                        if grid[group_start][col_idx].style:
                                            if not grid[group_start][col_idx].style.alignment:
                                                grid[group_start][col_idx].style.alignment = AlignmentStyle()
                                            grid[group_start][col_idx].style.alignment.horizontal = 'center'
                                            grid[group_start][col_idx].style.alignment.vertical = 'center'
                                        for clear_r in range(group_start + 1, group_end + 1):
                                            if clear_r in grid and col_idx in grid[clear_r]:
                                                grid[clear_r][col_idx].value = None
                            
                            group_start = r
                            group_value = current_value

        except Exception as fill_data_err:
            logger.error(f"Error during data filling loop: {fill_data_err}\n{traceback.format_exc()}")
            return None

        # Log completion summary
        logger.info(f"DataTableBuilder completed: {actual_rows_to_process} data rows generated")

        row_height = self.style_registry.get_row_height('data') if self.style_registry else None
        models = []
        for r in range(actual_rows_to_process):
            cells = sorted(list(grid[r].values()), key=lambda c: c.col_index)
            models.append(UnitRow(
                relative_index=r,
                height=row_height,
                cells=cells
            ))

        return models
    
    def _build_formula_string(self, formula_dict: Dict[str, Any], row_num: int) -> str:
        """
        Convert a formula dict to an Excel formula string.
        
        Args:
            formula_dict: Dict with 'template' and 'inputs' keys
            row_num: Current row number
        
        Returns:
            Excel formula string (e.g., "=B5*C5")
        """
        template = formula_dict.get('template', '')
        inputs = formula_dict.get('inputs', [])
        
        # Replace placeholders like {col_ref_0}, {col_ref_1}, etc.
        formula = template
        for i, input_id in enumerate(inputs):
            col_idx = self.col_id_map.get(input_id)
            if col_idx:
                col_letter = get_column_letter(col_idx)
                formula = formula.replace(f'{{col_ref_{i}}}', f'{col_letter}{{row}}')
        
        # Replace {row} with actual row number
        formula = formula.replace('{row}', str(row_num))
        
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = '=' + formula
        
        return formula

