import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter, column_index_from_string, coordinate_to_tuple

from ..models.footer import FooterData
from .json_template_builder import JsonTemplateStateBuilder
from ..models.layout import SheetLayoutState
from ..models.config.styling import SheetStylingModel
from ..models.config.layout import SheetLayoutModel
from ..mappers import ResolvedTableData, resolve_summary_payload
from .summary import SummaryBuilder
from ...system_config import sys_config


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
 
        logger.info(f"Layout built successfully for sheet '{self.sheet_name}'")
        return True

    def inject_images(self, stamp_cell: Optional[str] = None, signature_cell: Optional[str] = None):
        """
        Public method to inject template images after template footer restoration.
        """
        self._inject_images(stamp_cell=stamp_cell, signature_cell=signature_cell)

    def get_last_cell(self, col: Optional[str] = None) -> Tuple[int, str]:
        """
        Delegates to self.layout_state.get_last_cell().
        If col parameter is provided, overrides column_letter and updates cell_address.
        """
        max_row, col_letter = self.layout_state.get_last_cell()
        if col:
            col_letter = col.strip().upper()
        return max_row, col_letter

    def inject_image(self, image_path: str, cell_address: str = "N1") -> bool:
        """
        Inject an image at specified cell_address if image_path exists.
        Returns True on success, False on failure.
        """
        path = Path(image_path)
        if not path.is_file():
            logger.warning(f"Image path does not exist: {image_path}")
            return False
        try:
            img = Image(str(path))
            self.worksheet.add_image(img, cell_address)
            logger.info(f"Injected image '{image_path}' at {cell_address}")
            return True
        except Exception as e:
            logger.warning(f"Failed to inject image '{image_path}' at {cell_address}: {e}")
            return False

    def inject_image_with_offset(self, image_path: str, cell_address: str, col_offset_px: int = 0, row_offset_px: int = 0) -> bool:
        """
        Inject an image anchored at cell_address with pixel offsets.
        """
        path = Path(image_path)
        if not path.is_file():
            logger.warning(f"Image path does not exist: {image_path}")
            return False
        try:
            row_idx, col_idx = coordinate_to_tuple(cell_address)
        except Exception as e:
            logger.warning(f"Invalid cell address '{cell_address}': {e}")
            return False

        try:
            img = Image(str(path))
            col_off_emu = pixels_to_EMU(col_offset_px)
            row_off_emu = pixels_to_EMU(row_offset_px)
            w_px = getattr(img, 'width', None) or 100
            h_px = getattr(img, 'height', None) or 100
            img_w_emu = pixels_to_EMU(w_px)
            img_h_emu = pixels_to_EMU(h_px)

            col_marker = max(0, col_idx - 1)
            row_marker = max(0, row_idx - 1)
            marker = AnchorMarker(col=col_marker, colOff=col_off_emu, row=row_marker, rowOff=row_off_emu)
            size = XDRPositiveSize2D(img_w_emu, img_h_emu)
            img.anchor = OneCellAnchor(_from=marker, ext=size)
            self.worksheet.add_image(img)
            logger.info(f"Injected image '{image_path}' with offset ({col_offset_px}px, {row_offset_px}px) at {cell_address}")
            return True
        except Exception as e:
            logger.warning(f"Failed to inject image with offset '{image_path}' at {cell_address}: {e}")
            return False

    def _inject_images(self, stamp_cell: Optional[str] = None, signature_cell: Optional[str] = None):
        """
        Injects stamp/signature images from database/template_images and pre-loaded images using last cell position.
        """
        max_row, col_let = self.get_last_cell()
        stamp_address = stamp_cell or f"{col_let}{max_row}"

        stamp_path = sys_config.template_image_dir / "stamp.png"
        sig_path = sys_config.template_image_dir / "signiture.png"
        if not sig_path.exists():
            sig_path = sys_config.template_image_dir / "signature.png"

        stamp_width = 120
        stamp_height = 120
        if stamp_path.exists():
            try:
                stamp_img = Image(str(stamp_path))
                stamp_width = getattr(stamp_img, 'width', None) or 120
                stamp_height = getattr(stamp_img, 'height', None) or 120
            except Exception:
                pass

        self.inject_image(str(stamp_path), stamp_address)

        if signature_cell is not None:
            self.inject_image(str(sig_path), signature_cell)
        else:
            col_off_px = int(stamp_width * 0.35)
            row_off_px = int(stamp_height * 0.25)
            self.inject_image_with_offset(str(sig_path), stamp_address, col_offset_px=col_off_px, row_offset_px=row_off_px)

        if self.pre_loaded_images:
            logger.info(f"Injecting {len(self.pre_loaded_images)} pre-loaded images into sheet '{self.sheet_name}'")
            for idx, img_item in enumerate(self.pre_loaded_images):
                try:
                    inc_offset = idx * 10
                    if isinstance(img_item, (str, Path)):
                        self.inject_image_with_offset(str(img_item), stamp_address, col_offset_px=inc_offset, row_offset_px=inc_offset)
                    elif isinstance(img_item, Image):
                        try:
                            row_idx, col_idx = coordinate_to_tuple(stamp_address)
                            col_off_emu = pixels_to_EMU(inc_offset)
                            row_off_emu = pixels_to_EMU(inc_offset)
                            w_px = getattr(img_item, 'width', None) or 100
                            h_px = getattr(img_item, 'height', None) or 100
                            img_w_emu = pixels_to_EMU(w_px)
                            img_h_emu = pixels_to_EMU(h_px)
                            marker = AnchorMarker(col=max(0, col_idx - 1), colOff=col_off_emu, row=max(0, row_idx - 1), rowOff=row_off_emu)
                            size = XDRPositiveSize2D(img_w_emu, img_h_emu)
                            img_item.anchor = OneCellAnchor(_from=marker, ext=size)
                            self.worksheet.add_image(img_item)
                        except Exception:
                            self.worksheet.add_image(img_item, stamp_address)
                    else:
                        self.worksheet.add_image(img_item, stamp_address)
                    logger.debug(f"Injected pre-loaded image at {stamp_address}")
                except Exception as e:
                    logger.warning(f"Failed to inject pre-loaded image into worksheet: {e}")
