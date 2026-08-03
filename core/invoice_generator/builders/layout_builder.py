import logging
from typing import Any, Dict, List, Optional

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

from ..models.footer import FooterData
from .json_template_builder import JsonTemplateStateBuilder
from ..models.layout import SheetLayoutState
from ..models.config.styling import SheetStylingModel
from ..models.config.layout import SheetLayoutModel
from ..mappers import ResolvedTableData, resolve_summary_payload
from .summary import SummaryBuilder


# Initialize logger for this module
logger = logging.getLogger(__name__)


class LayoutBuilder:
    """
    The Director in the Builder pattern.
    Coordinates all builders to construct the complete document layout.
    """
    def __init__(
        self,
        workbook: Workbook,
        worksheet: Worksheet,
        template_worksheet: Worksheet,
        sheet_styling: SheetStylingModel,
        sheet_layout: SheetLayoutModel,
        resolved_data: ResolvedTableData,
        sheet_name: str,
        all_sheet_configs: Dict[str, Any],
        args: Any = None,
        total_net_weight: Optional[float] = None,
        total_gross_weight: Optional[float] = None,
        skip_template_header_restoration: bool = False,
        skip_header_builder: bool = False,
        skip_data_table_builder: bool = False,
        skip_footer_builder: bool = False,
        skip_template_footer_restoration: bool = False,
        template_state_builder: Optional[JsonTemplateStateBuilder] = None,
        template_json_config: Optional[Dict[str, Any]] = None,
        layout_state: Optional[SheetLayoutState] = None,
        pre_loaded_images: Optional[List[Image]] = None,
        invoice_data: Optional[Dict[str, Any]] = None,
        table_key: Optional[str] = None
    ):
        """
        Initialize LayoutBuilder with strict model architecture.
        """
        self.workbook = workbook
        self.worksheet = worksheet
        self.template_worksheet = template_worksheet
        self.sheet_styling = sheet_styling
        self.sheet_layout = sheet_layout
        self.resolved_data = resolved_data
        self.sheet_name = sheet_name
        self.all_sheet_configs = all_sheet_configs
        self.args = args
        self.total_net_weight = total_net_weight
        self.total_gross_weight = total_gross_weight
        
        self.skip_template_header_restoration = skip_template_header_restoration
        self.skip_header_builder = skip_header_builder
        self.skip_data_table_builder = skip_data_table_builder
        self.skip_footer_builder = skip_footer_builder
        self.skip_template_footer_restoration = skip_template_footer_restoration
        
        self.template_state_builder = template_state_builder
        self.template_json_config = template_json_config
        self.layout_state = layout_state or SheetLayoutState()
        self.pre_loaded_images = pre_loaded_images or []
        self.invoice_data = invoice_data
        self.table_key = table_key

        # Store results after build
        self.next_row_after_footer = -1
        self.data_start_row = -1
        self.data_end_row = -1
        self.footer_data: Optional[FooterData] = None
        self.leather_summary: Optional[Dict[str, Any]] = None
        
        logger.debug(f"LayoutBuilder initialized for '{self.sheet_name}' with models")

    def build(self) -> bool:
        """
        Orchestrates all builders in the correct sequence.
        """
        logger.info(f"Building layout for sheet '{self.sheet_name}'")
        logger.debug(f"Reading from template, writing to output worksheet")
        
        # Calculate header boundaries for template state capture
        header_row = self.sheet_layout.structure.header_row
        
        if self.layout_state:
            # Use top-level layout state allocator if available
            table_header_row = self.layout_state.next_free_row
            logger.info(f"Using layout_state.next_free_row for table_header_row: {table_header_row}")
        else:
            table_header_row = header_row

        logger.debug(f"[LayoutBuilder DEBUG] sheet_name={self.sheet_name}, header_row={header_row}, table_header_row={table_header_row}")
        
        # 3. Template State Capture
        if self.template_state_builder:
            logger.info(f"Using pre-captured template state (multi-table optimization)")
        elif self.template_json_config and self.sheet_name in self.template_json_config:
            logger.info(f"Using JSON-based template state for sheet '{self.sheet_name}'")
            try:
                sheet_layout_json = self.template_json_config[self.sheet_name]
                self.template_state_builder = JsonTemplateStateBuilder(
                    sheet_layout_data=sheet_layout_json
                )
                logger.info(f"JSON Template loaded: Header ends {self.template_state_builder.header_end_row}, Footer starts {self.template_state_builder.template_footer_start_row}")
                
            except Exception as e:
                logger.critical(f"CRITICAL: JsonTemplateStateBuilder failed for '{self.sheet_name}': {e}", exc_info=True)
                return False
        else:
            logger.critical(f"CRITICAL: No JSON template found for sheet '{self.sheet_name}'.")
            return False
            
        # 4. Table Builder delegation
        from .table import TableBuilder
        from .table.builder import TableBuilderConfig
        import copy
        
        # Resolve active columns/mappings based on DAF and custom mode
        daf_mode = getattr(self.args, 'DAF', False) if self.args else False
        custom_mode = getattr(self.args, 'custom', False) if self.args else False
        
        columns_source = self.sheet_layout.structure.columns
        if self.skip_header_builder:
            columns_source = []
            
        from ..models.config.layout import StructureConfigModel
        structure = StructureConfigModel(
            header_row=self.sheet_layout.structure.header_row,
            columns=columns_source
        )
        
        bundled_columns, column_index_mapping, column_mapping, column_colspan = (
            structure.resolve_mappings(DAF_mode=daf_mode, custom_mode=custom_mode)
        )
        
        resolved_data = self.resolved_data
        if self.skip_data_table_builder:
            from ..mappers.models import ResolvedTableData
            resolved_data = ResolvedTableData(data_rows=[])
            
        sheet_layout = copy.deepcopy(self.sheet_layout)
        if self.skip_footer_builder:
            sheet_layout.footer = None
            
         # Attach resolved mapping properties directly to sheet_layout
        sheet_layout.bundled_columns = bundled_columns
        sheet_layout.column_index_mapping = column_index_mapping
        sheet_layout.column_mapping = column_mapping
        sheet_layout.column_colspan = column_colspan
        
        # Pre-calculate vertical merges using the transformer
        from ..utils.merge_transformer import apply_vertical_merges
        resolved_data.data_rows = apply_vertical_merges(resolved_data.data_rows)
            
        if not resolved_data.footer:
            from ..mappers.models import ResolvedTableFooter
            resolved_data.footer = ResolvedTableFooter()
            
        if not resolved_data.footer.weight_summary:
            resolved_data.footer.weight_summary = {
                'net': self.total_net_weight or 0.0,
                'gross': self.total_gross_weight or 0.0
            }
        if self.invoice_data and 'footer_data' in self.invoice_data:
            footer_dict = self.invoice_data['footer_data']
            table_total = None
            if self.table_key is not None and 'table_totals' in footer_dict:
                table_totals = footer_dict['table_totals']
                if isinstance(table_totals, list):
                    try:
                        idx = int(self.table_key)
                        if 0 <= idx < len(table_totals):
                            table_total = table_totals[idx]
                    except ValueError:
                        pass
                elif isinstance(table_totals, dict) and str(self.table_key) in table_totals:
                    table_total = table_totals[str(self.table_key)]

            resolved_data.footer.grand_total = table_total or footer_dict.get('grand_total', {})
            resolved_data.footer.leather_summary = footer_dict.get('leather_summary', [])
            
        sheet_styling = copy.deepcopy(self.sheet_styling)
        if daf_mode and sheet_styling:
            if hasattr(sheet_styling, "columns") and "col_unit_price" in sheet_styling.columns:
                sheet_styling.columns["col_unit_price"].format = "#,##0.0000000"
            elif isinstance(sheet_styling, dict) and "columns" in sheet_styling and "col_unit_price" in sheet_styling["columns"]:
                if isinstance(sheet_styling["columns"]["col_unit_price"], dict):
                    sheet_styling["columns"]["col_unit_price"]["format"] = "#,##0.0000000"

        config = TableBuilderConfig(
            worksheet=self.worksheet,
            sheet_styling=sheet_styling,
            sheet_layout=sheet_layout,
            resolved_data=resolved_data,
            table_key=self.table_key
        )
        table_builder = TableBuilder(config=config)
        
        success = table_builder.build(start_row=table_header_row, layout_state=self.layout_state)
        if not success:
            logger.error("Table building failed")
            return False
            
        # Copy properties for downstream components and compatibility
        self.grid = table_builder.grid
        self.footer_data = table_builder.footer_data
        self.data_start_row = table_builder.data_start_row
        self.data_end_row = table_builder.data_end_row
        self.next_row_after_footer = table_builder.next_row_after_footer
        self.column_index_mapping = getattr(table_builder, 'column_index_mapping', {})
        
        # 5b. NOW restore template header - AFTER table is built
        if not self.skip_template_header_restoration:
            logger.info(f"Restoring template header AFTER table build (correct column alignment)")
            try:
                actual_num_cols = self.grid.num_columns
                
                gen_mode = "standard"
                if self.args:
                    if getattr(self.args, 'DAF', False): gen_mode = "daf"
                    elif getattr(self.args, 'custom', False): gen_mode = "custom"
                
                if self.template_state_builder:
                    self.template_state_builder.restore_header_only(
                        target_worksheet=self.worksheet,
                        actual_num_cols=actual_num_cols,
                        mode=gen_mode,
                        layout_state=self.layout_state,
                        column_index_mapping=self.column_index_mapping
                    )
                    logger.info(f"Template header restored successfully with {actual_num_cols} columns")
            except Exception as e:
                logger.error(f"Failed to restore template header after table build: {e}", exc_info=True)
                return False
        else:
            logger.debug(f"Skipping template header restoration")
        
        # 6b. Apply static column widths
        logger.info("Applying static config widths (auto-fit disabled)")
        try:
            col_id_to_idx = self.grid.column_mapping
            widths = {}
            columns = self.sheet_layout.structure.columns
            
            def extract_widths(cols):
                for col in cols:
                    col_id = col.id
                    col_w = col.width
                    if col_id and col_w is not None:
                        widths[col_id] = col_w
                    if col.children:
                        extract_widths(col.children)
                        
            extract_widths(columns)
            
            for col_id, col_idx in col_id_to_idx.items():
                width = widths.get(col_id)
                if width:
                    col_letter = get_column_letter(col_idx)
                    self.worksheet.column_dimensions[col_letter].width = float(width)
                    logger.debug(f"Applied static width {width} to {col_id} ({col_letter})")
        except Exception as e:
            logger.error(f"Failed to apply static column widths: {e}", exc_info=True)
 
        # 7. Inject Template Images
        self._inject_images()
        
        logger.info(f"Layout built successfully for sheet '{self.sheet_name}'")
        return True

    def _inject_images(self):
        """
        Injects pre-loaded images into the worksheet.
        """
        if not self.pre_loaded_images:
            logger.debug("No pre-loaded images to inject")
            return
            
        logger.info(f"Injecting {len(self.pre_loaded_images)} pre-loaded images")
        for img in self.pre_loaded_images:
            try:
                # Default placement at N1 (as requested)
                self.worksheet.add_image(img, 'N1')
                logger.debug(f"Injected image at N1")
            except Exception as e:
                logger.warning(f"Failed to inject pre-loaded image: {e}")
