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
from ..models.table_adapter import ResolvedTableData

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
        final_grand_total_pallets: int = 0,
        total_net_weight: Optional[float] = None,
        total_gross_weight: Optional[float] = None,
        is_last_table: bool = False,
        show_grand_total_addons: bool = False,
        skip_template_header_restoration: bool = False,
        skip_header_builder: bool = False,
        skip_data_table_builder: bool = False,
        skip_footer_builder: bool = False,
        skip_template_footer_restoration: bool = False,
        template_state_builder: Optional[JsonTemplateStateBuilder] = None,
        template_json_config: Optional[Dict[str, Any]] = None,
        layout_state: Optional[SheetLayoutState] = None,
        pre_loaded_images: Optional[List[Image]] = None
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
        self.final_grand_total_pallets = final_grand_total_pallets
        self.total_net_weight = total_net_weight
        self.total_gross_weight = total_gross_weight
        self.is_last_table = is_last_table
        self.show_grand_total_addons = show_grand_total_addons
        
        self.skip_template_header_restoration = skip_template_header_restoration
        self.skip_header_builder = skip_header_builder
        self.skip_data_table_builder = skip_data_table_builder
        self.skip_footer_builder = skip_footer_builder
        self.skip_template_footer_restoration = skip_template_footer_restoration
        
        self.template_state_builder = template_state_builder
        self.template_json_config = template_json_config
        self.layout_state = layout_state or SheetLayoutState()
        self.pre_loaded_images = pre_loaded_images or []

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
        table_builder = TableBuilder(
            workbook=self.workbook,
            worksheet=self.worksheet,
            sheet_styling=self.sheet_styling,
            sheet_layout=self.sheet_layout,
            resolved_data=self.resolved_data,
            sheet_name=self.sheet_name,
            args=self.args,
            final_grand_total_pallets=self.final_grand_total_pallets,
            total_net_weight=self.total_net_weight,
            total_gross_weight=self.total_gross_weight,
            is_last_table=self.is_last_table,
            show_grand_total_addons=self.show_grand_total_addons,
            skip_header_builder=self.skip_header_builder,
            skip_data_table_builder=self.skip_data_table_builder,
            skip_footer_builder=self.skip_footer_builder,
            layout_state=self.layout_state
        )
        
        success = table_builder.build(start_row=table_header_row)
        if not success:
            logger.error("Table building failed")
            return False
            
        # Copy properties for downstream components and compatibility
        self.grid = table_builder.grid
        self.footer_data = table_builder.footer_data
        self.data_start_row = table_builder.data_start_row
        self.data_end_row = table_builder.data_end_row
        self.next_row_after_footer = table_builder.next_row_after_footer
        
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
                        layout_state=self.layout_state
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
 
        # 7. Template Footer Restoration
        if self.template_state_builder and not self.skip_template_footer_restoration:
            try:
                actual_num_cols = self.grid.num_columns
                
                if self.is_last_table:
                    logger.info(f"--- RESTORING TEMPLATE FOOTER (Last Table) ---")
                    
                    gen_mode = "standard"
                    if self.args:
                        if getattr(self.args, 'DAF', False): gen_mode = "daf"
                        elif getattr(self.args, 'custom', False): gen_mode = "custom"

                    self.template_state_builder.restore_template_footer(
                        target_worksheet=self.worksheet,
                        footer_start_row=self.next_row_after_footer,
                        actual_num_cols=actual_num_cols,
                        mode=gen_mode,
                        layout_state=self.layout_state
                    )
                else:
                    logger.info(f"Skipping template footer restoration (Not last table)")
                logger.info(f"Template footer restored successfully")
            except Exception as e:
                logger.error(f"Failed to restore template footer: {e}", exc_info=True)
        else:
            logger.debug("Skipping template footer restoration")

        # 8. Inject Template Images
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
