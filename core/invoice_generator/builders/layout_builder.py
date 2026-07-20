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
from ..mappers.models import ResolvedTableData

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
        is_last_table: bool = False,
        skip_template_header_restoration: bool = False,
        skip_header_builder: bool = False,
        skip_data_table_builder: bool = False,
        skip_footer_builder: bool = False,
        skip_template_footer_restoration: bool = False,
        template_state_builder: Optional[JsonTemplateStateBuilder] = None,
        template_json_config: Optional[Dict[str, Any]] = None,
        layout_state: Optional[SheetLayoutState] = None,
        pre_loaded_images: Optional[List[Image]] = None,
        invoice_data: Optional[Dict[str, Any]] = None
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
        self.is_last_table = is_last_table
        
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
            resolved_data.footer.grand_total = footer_dict.get('grand_total', {})
            resolved_data.footer.leather_summary = footer_dict.get('leather_summary', [])
            
        config = TableBuilderConfig(
            worksheet=self.worksheet,
            sheet_styling=self.sheet_styling,
            sheet_layout=sheet_layout,
            resolved_data=resolved_data
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
 
        # 6c. Build Page-level Summary
        if self.is_last_table and self.sheet_layout.summary and self.sheet_layout.summary.rows:
            logger.info("Building page-level summary section")
            try:
                # Prepare payload
                payload = {}
                if self.invoice_data and 'footer_data' in self.invoice_data:
                    footer_data_dict = self.invoice_data['footer_data']
                    payload.update(footer_data_dict.get('grand_total', {}))
                    payload['leather_summary'] = footer_data_dict.get('leather_summary', [])
                
                # Fetch pallet count from footer_data or default
                pallet_count = 0
                if self.invoice_data and 'footer_data' in self.invoice_data:
                    grand_total = self.invoice_data['footer_data'].get('grand_total', {})
                    pallet_count = grand_total.get('col_pallet_count', grand_total.get('pallet_count', 0))
                
                # Fallback to self.footer_data pallet count
                if not pallet_count and self.footer_data:
                    pallet_count = self.footer_data.total_pallets
                
                payload.setdefault('pallet_count', int(pallet_count))
                payload.setdefault('multiple', "S" if payload['pallet_count'] != 1 else "")
                payload.setdefault('weight_net', payload.get('col_net', 0.0))
                payload.setdefault('weight_gross', payload.get('col_gross', 0.0))
                payload.setdefault('leather_summary', [])

                from .table.summary import SummaryBuilder
                summary_builder = SummaryBuilder(
                    worksheet=self.worksheet,
                    summary_config=self.sheet_layout.summary,
                    payload=payload,
                    start_row=self.next_row_after_footer,
                    last_grid=self.grid
                )
                self.next_row_after_footer = summary_builder.build()
            except Exception as e:
                logger.error(f"[LayoutBuilder] SummaryBuilder failed: {e}", exc_info=True)
                return False

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
                        layout_state=self.layout_state,
                        column_index_mapping=self.column_index_mapping
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
