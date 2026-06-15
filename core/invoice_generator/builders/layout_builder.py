import logging
from typing import Any, Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

from ..styling.models import StylingConfigModel, FooterData
from ..data.table_calculator import TableCalculator
from .json_template_builder import JsonTemplateStateBuilder
from openpyxl.drawing.image import Image
from ...system_config import sys_config, ConfigurationError
from ..models.layout import SheetLayoutState
from ..styling.style_registry import StyleRegistry
from openpyxl.utils import get_column_letter

# Initialize logger for this module
logger = logging.getLogger(__name__)

class LayoutBuilder:
    """
    The Director in the Builder pattern.
    Coordinates all builders to construct the complete document layout.
    
    RECOMMENDED USAGE (Modern Bundled Config Approach):
        Use BuilderConfigResolver to prepare configuration bundles, then pass them
        via style_config, context_config, and layout_config parameters. This approach
        centralizes config resolution logic and eliminates duplication.
        
        Example:
            from invoice_generator.config.builder_config_resolver import BuilderConfigResolver
            
            resolver = BuilderConfigResolver(
                config_loader=config_loader,
                sheet_name='Invoice',
                worksheet=worksheet,
                args=args,
                invoice_data=invoice_data,
                pallets=31
            )
            
            # Get bundles - resolver handles all data extraction
            style_config, context_config, layout_config, data_config = resolver.get_datatable_bundles()
            
            layout_builder = LayoutBuilder(
                workbook=workbook,
                worksheet=worksheet,
                template_worksheet=template,
                style_config=style_config,
                context_config=context_config,
        """
    def __init__(
        self,
        workbook: Workbook,
        worksheet: Worksheet,
        template_worksheet: Worksheet,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Dict[str, Any],
        template_state_builder: Optional[JsonTemplateStateBuilder] = None,
        template_json_config: Optional[Dict[str, Any]] = None,
        layout_state: Optional[SheetLayoutState] = None
    ):
        """
        Initialize LayoutBuilder with strict bundle architecture.
        
        Args:
            workbook: Output workbook (writable)
            worksheet: Output worksheet (writable)
            template_worksheet: Template worksheet (read-only)
            style_config: Bundle containing styling configuration
            context_config: Bundle containing context (sheet_name, invoice_data, args, etc.)
            layout_config: Bundle containing layout rules, structure, and resolved data
            template_state_builder: Optional pre-captured template state (optimization)
            template_json_config: Optional template JSON config
            layout_state: Optional layout state tracking occupied rows
        """
        self.workbook = workbook
        self.worksheet = worksheet
        self.template_worksheet = template_worksheet
        if not layout_state:
            layout_state = SheetLayoutState()
        self.layout_state = layout_state
        
        # Unpack Style Bundle
        self.styling_config = style_config.get('styling_config')
        
        # Unpack Context Bundle
        self.sheet_name = context_config.get('sheet_name')
        self.invoice_data = context_config.get('invoice_data')
        self.all_sheet_configs = context_config.get('all_sheet_configs')
        self.args = context_config.get('args')
        self.final_grand_total_pallets = context_config.get('final_grand_total_pallets', 0)
        self.total_net_weight = context_config.get('total_net_weight')
        self.total_gross_weight = context_config.get('total_gross_weight')
        self.is_last_table = context_config.get('is_last_table', False)
        self.show_grand_total_addons = context_config.get('show_grand_total_addons', False)
        
        # Unpack Layout Bundle
        self.sheet_config = layout_config.get('sheet_config', {})
        
        # Skip flags
        self.skip_template_header_restoration = layout_config.get('skip_template_header_restoration', False)
        self.skip_header_builder = layout_config.get('skip_header_builder', False)
        self.skip_data_table_builder = layout_config.get('skip_data_table_builder', False)
        self.style_config = style_config or {}
        self.context_config = context_config or {}
        self.layout_config = layout_config or {}
        self.template_state_builder = template_state_builder
        self.skip_footer_builder = self.layout_config.get('skip_footer_builder', False)
        
        # We need this to apply padding/dimensions after build
        self.header_info = self.layout_config.get('header_info', {})
        self.skip_template_footer_restoration = layout_config.get('skip_template_footer_restoration', False)
        
        # Data Source (Must be provided via resolved_data in layout_config)
        self.provided_resolved_data = layout_config.get('resolved_data')
        self.provided_header_info = layout_config.get('header_info')
        self.provided_mapping_rules = layout_config.get('mapping_rules')
        
        # Pre-captured template state
        self.pre_captured_template_state = template_state_builder
        self.template_json_config = template_json_config
        
        logger.debug(f"LayoutBuilder initialized for '{self.sheet_name}' with pure bundle config")
        
        # Store results after build
        self.header_info = None
        self.next_row_after_footer = -1
        self.data_start_row = -1
        self.data_end_row = -1
        self.template_state_builder: Optional[JsonTemplateStateBuilder] = None
        self.footer_data: Optional[FooterData] = None
        self.leather_summary: Optional[Dict[str, Any]] = None

    def build(self) -> bool:
        """
        Orchestrates all builders in the correct sequence.
        Reads template state from template_worksheet, writes to self.worksheet (output).
        This completely avoids merge conflicts since template and output are separate.
        """
        logger.info(f"Building layout for sheet '{self.sheet_name}'")
        logger.debug(f"Reading from template, writing to output worksheet")
        
        # 1. Text Replacement (if enabled) - Pre-processing
        # Removed per user request
        
        # 2. Calculate header boundaries for template state capture
        structure = self.sheet_config.get('structure', {})
        header_row = structure.get('header_row') # Correct placement for header offset

        # IMPORTANT: Clarify terminology - there are TWO types of headers:
        # 1. TEMPLATE HEADER: Decorative header section (company name, logo, etc.) - rows 1 to (table_header_row - 1)
        # 2. TABLE HEADER: Column headers for data table (e.g., "Item", "Quantity", "Price") - at table_header_row
        
        # Get table_header_row from config (where the data table column headers are)
        # For multi-table sheets, multi_table_processor dynamically injects the correct
        # expected header_row into self.sheet_config ['structure']['header_row'].
        # We MUST respect this injected value over the static global sheet_layout original value.
        sheet_layout = self.all_sheet_configs.get(self.sheet_name, {}) if self.all_sheet_configs else {}
        
        if self.layout_state:
            # Use top-level layout state allocator if available
            table_header_row = self.layout_state.next_free_row
            logger.info(f"Using layout_state.next_free_row for table_header_row: {table_header_row}")
        # Priority 1: Injected structure.header_row from multi_table_processor
        elif self.sheet_config and 'structure' in self.sheet_config and 'header_row' in self.sheet_config['structure']:
            table_header_row = self.sheet_config['structure']['header_row']
        # Priority 2: Original static template header_row
        else:
            table_header_row = sheet_layout.get('structure', {}).get('header_row', header_row)
            
        if table_header_row is None:
            raise ConfigurationError(f"CRITICAL: No 'header_row' found for sheet '{self.sheet_name}'. Check configuration structure.")

        header_row_for_builder = table_header_row
        logger.debug(f"[LayoutBuilder DEBUG] sheet_name={self.sheet_name}, header_row={header_row}, table_header_row={table_header_row}")

        logger.debug(f"[LayoutBuilder DEBUG] all_sheet_configs keys: {list(self.all_sheet_configs.keys()) if self.all_sheet_configs else 'None'}")
        
        # Template decorative header spans from row 1 to the row BEFORE the table header
        template_header_start_row = 1
        template_header_end_row = table_header_row - 1  # Decorative header ends BEFORE table header
        
        # Calculate footer_start_row from template (estimate: table_header_row + 2-row table header + minimal data rows)
        # Table header is at table_header_row, second header row at table_header_row + 1
        # Data starts at table_header_row + 2, footer would be around data_start + 2 rows
        # Calculate footer_start_row dynamically from template
        # 3. Template State Capture
        if self.pre_captured_template_state:
            logger.info(f"Using pre-captured template state (multi-table optimization)")
            self.template_state_builder = self.pre_captured_template_state
            logger.debug(f"Reusing template state")
        elif self.template_json_config and self.sheet_name in self.template_json_config:
            # === NEW JSON-BASED PATH ===
            logger.info(f"Using JSON-based template state for sheet '{self.sheet_name}'")
            try:
                sheet_layout_json = self.template_json_config[self.sheet_name]
                self.template_state_builder = JsonTemplateStateBuilder(
                    sheet_layout_data=sheet_layout_json
                )
                
                # Setup critical boundaries from the loaded builder
                template_header_end_row = self.template_state_builder.header_end_row
                template_footer_start_row = self.template_state_builder.template_footer_start_row
                
                logger.info(f"JSON Template loaded: Header ends {template_header_end_row}, Footer starts {template_footer_start_row}")
                
            except Exception as e:
                logger.critical(f"CRITICAL: JsonTemplateStateBuilder failed for '{self.sheet_name}': {e}", exc_info=True)
                return False
        else:
            # JSON template required - XLSX scanning has been removed
            logger.critical(f"CRITICAL: No JSON template found for sheet '{self.sheet_name}'. XLSX scanning has been removed.")
            return False
            
        # Common: Text replacements removed per user request
        
        # 3b. Template header restoration DEFERRED - will be done AFTER table building
        # This ensures template content aligns with actual column count after filtering
        logger.debug(f"Deferring template header restoration until after table building")
        
        # 4. Table Builder delegation
        from .table import TableBuilder
        table_builder = TableBuilder(
            workbook=self.workbook,
            worksheet=self.worksheet,
            style_config=self.style_config,
            context_config=self.context_config,
            layout_config=self.layout_config,
            layout_state=self.layout_state
        )
        
        success = table_builder.build(start_row=table_header_row)
        if not success:
            logger.error("Table building failed")
            return False
            
        # Copy properties for downstream components and compatibility
        self.header_info = table_builder.header_info
        self.footer_data = table_builder.footer_data
        self.data_start_row = table_builder.data_start_row
        self.data_end_row = table_builder.data_end_row
        self.next_row_after_footer = table_builder.next_row_after_footer
        
        # 5b. NOW restore template header - AFTER table is built
        # This ensures template content aligns with actual number of columns used
        # CRITICAL: This should only restore decorative header (rows 1 to table_header_row-1)
        # It must NOT overwrite the table header row that HeaderBuilder styled
        if not self.skip_template_header_restoration:
            logger.info(f"Restoring template header AFTER table build (correct column alignment)")
            try:
                # Get actual column count from header_info (this reflects filtered columns)
                actual_num_cols = self.header_info.get('num_columns', None)
                table_header_row_num = self.header_info.get('second_row_index', 0)
                logger.debug(f"Template header will use actual column count: {actual_num_cols}")
                if self.template_state_builder:
                    logger.debug(f"Template header ends at row {self.template_state_builder.header_end_row}")
                logger.debug(f"Table header row is at: {table_header_row_num}")
                logger.debug(f"These should NOT overlap! (template_end < table_header)")
                # DO NOT apply column mapping to the template header!
                # The user specifically requested that we do not skip anything
                # when capturing/restoring the template wrapper.
                
                # Resolve generation mode for mode-dependent header values
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
                    logger.info(f"Template header restored successfully with {actual_num_cols} columns (rows 1-{self.template_state_builder.header_end_row})")
            except Exception as e:
                logger.error(f"Failed to restore template header after table build")
                logger.error(f"Error: {e}", exc_info=True)
                return False
        else:
            logger.debug(f"Skipping template header restoration (skip_template_header_restoration=True)")
        
        # 6b. Apply static column widths (auto_fit completely removed)
        logger.info("Applying static config widths (auto-fit disabled)")
        try:
            # Map column IDs to their physical column indices
            col_id_to_idx = self.header_info.get('column_id_map', {})
            
            # Apply static widths from the styling registry
            if hasattr(self, 'styling_config') and self.styling_config:
                # New style registry format stores widths in the columns dict
                columns_config = self.styling_config.get('columns', {}) if isinstance(self.styling_config, dict) else {}
                for col_id, col_idx in col_id_to_idx.items():
                    col_cfg = columns_config.get(col_id, {})
                    width = col_cfg.get('width')
                    if width:
                        col_letter = get_column_letter(col_idx)
                        self.worksheet.column_dimensions[col_letter].width = float(width)
                        logger.debug(f"Applied static width {width} to {col_id} ({col_letter})")
        except Exception as e:
            logger.error(f"Failed to apply static column widths: {e}", exc_info=True)

        # 7. Template Footer Restoration
        # This restores the static content (signatures, etc.) from the JSON template
        # that appears AFTER the dynamic table footer.
        skip_template_footer = self.layout_config.get('skip_template_footer_restoration', False)
        
        if self.template_state_builder and not skip_template_footer:
            try:
                # Get actual column count if not already set
                actual_num_cols = self.header_info.get('num_columns', None)
                
                # CRITICAL FIX: Only restore template footer if this is the LAST table on the sheet.
                # Otherwise, the static footer content (signatures, etc.) will be printed in the middle
                # of the sheet, distorting subsequent tables.
                if self.is_last_table:
                    logger.info(f"--- RESTORING TEMPLATE FOOTER (Last Table) ---")
                    logger.info(f"next_row_after_footer: {self.next_row_after_footer}")
                    
                    # Resolve generation mode for mode-dependent footer values
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
            logger.debug("Skipping template footer restoration (no template_state_builder)")

        # 8. Inject Template Images (New Feature)
        self._inject_images()
        
        logger.info(f"Layout built successfully for sheet '{self.sheet_name}'")
        
        return True

    def _inject_images(self):
        """
        Injects images from the configured directory into the worksheet.
        """
        try:
            img_dir = sys_config.template_image_dir
            if not img_dir.exists():
                logger.debug(f"Template image directory not found: {img_dir}")
                return

            images = list(img_dir.glob("*"))
            if not images:
                logger.debug(f"No images found in {img_dir}")
                return

            logger.info(f"Injecting {len(images)} images from {img_dir}")
            
            for i, img_path in enumerate(images):
                if img_path.suffix.lower() not in ['.png', '.jpg', '.jpeg', '.bmp', '.gif']:
                    continue
                    
                try:
                    img = Image(str(img_path))
                    
                    # Resize to 70px height (maintaining aspect ratio)
                    # This only affects display size; original image data is preserved
                    target_height = 140
                    if img.height > 0:
                        aspect_ratio = img.width / img.height
                        new_width = target_height * aspect_ratio
                        
                        img.height = target_height
                        img.width = new_width
                        
                    # Default placement at N1 (as requested)
                    self.worksheet.add_image(img, 'N1')
                    logger.debug(f"Injected image: {img_path.name} (resized to 70px height) at N1")
                except Exception as e:
                    logger.warning(f"Failed to inject image {img_path.name}: {e}")
        except Exception as e:
            logger.error(f"Image injection failed: {e}", exc_info=True)
    

