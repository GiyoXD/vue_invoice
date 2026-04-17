# invoice_generator/processors/base_processor.py
from abc import ABC, abstractmethod
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import argparse
from typing import Dict, Any, Optional
from core.system_config import ConfigurationError

class SheetProcessor(ABC):
    """
    Abstract base class for processing a single worksheet in an invoice workbook.
    Defines the common interface for all concrete processor implementations.
    """
    def __init__(
        self,
        template_workbook: Workbook,
        output_workbook: Workbook,
        template_worksheet: Worksheet,
        output_worksheet: Worksheet,
        sheet_name: str,
        sheet_config: Dict[str, Any],
        data_source_indicator: str,
        invoice_data: Dict[str, Any],
        cli_args: argparse.Namespace,
        final_grand_total_pallets: int,
        config_loader: Optional[Any] = None,
        data_mapping_config: Optional[Dict[str, Any]] = None  # Deprecated, use config_loader instead
    ):
        """
        Initializes the processor with all necessary data and configurations.

        Args:
            template_workbook: The template workbook (READ-ONLY usage for state capture)
            output_workbook: The output workbook (WRITABLE for final output)
            template_worksheet: The template worksheet to read state from
            output_worksheet: The output worksheet to write to
            sheet_name: The name of the worksheet.
            sheet_config: The specific configuration section for this sheet.
            data_mapping_config: The entire 'data_mapping' section from the config.
            data_source_indicator: The key indicating which data source to use.
            invoice_data: The complete input data dictionary.
            cli_args: The command-line arguments.
            final_grand_total_pallets: The pre-calculated total number of pallets.
        """
        self.template_workbook = template_workbook
        self.output_workbook = output_workbook
        self.template_worksheet = template_worksheet
        self.output_worksheet = output_worksheet
        
        # Keep old names for backward compatibility during transition
        self.workbook = output_workbook
        self.worksheet = output_worksheet
        
        self.sheet_name = sheet_name
        self.sheet_config = sheet_config
        self.data_mapping_config = data_mapping_config
        self.data_source_indicator = data_source_indicator
        self.invoice_data = invoice_data
        self.args = cli_args
        self.final_grand_total_pallets = final_grand_total_pallets
        self.config_loader = config_loader  # Store config loader for resolver usage
        self.processing_successful = True
        
        # New: Store config loader for direct bundled config access
        self.config_loader = config_loader
        self._use_bundled = config_loader is not None
        
        # New: Strict Header Row Validation (Mandatory for all table-based sheets)
        self.layout_config = self.sheet_config.get('layout_config', {}) if self.sheet_config else {}
        structure = self.layout_config.get('structure', {})
        self.header_row = structure.get('header_row')
        
        if self.header_row is None:
             raise ConfigurationError(f"CRITICAL: No 'header_row' found in configuration for sheet '{self.sheet_name}'. "
                                    f"Please ensure layout_config -> structure -> header_row is defined in your JSON.")

        # --- GLOBAL UNIQUENESS SCAN (Invoice Scope) ---
        # Calculate once in BaseProcessor so all children share the same authority.
        self.all_global_descriptions = set()
        
        # 1. Scan raw data
        if self.invoice_data:
            # Check multi_table
            multi_table = self.invoice_data.get('multi_table', [])
            if isinstance(multi_table, list):
                for table in multi_table:
                    if isinstance(table, list):
                        for row in table:
                            d = str(row.get('col_desc', "")).strip()
                            if d: self.all_global_descriptions.add(d)
            
            # Check single_table
            single_table = self.invoice_data.get('single_table', {})
            if isinstance(single_table, dict):
                for agg_key in ['aggregation', 'aggregation_custom', 'aggregation_DAF']:
                    agg_data = single_table.get(agg_key, [])
                    if isinstance(agg_data, list):
                        for row in agg_data:
                            d = str(row.get('col_desc', "")).strip()
                            if d: self.all_global_descriptions.add(d)

        # 2. Scan Fallbacks in Configuration (Truth if data is empty)
        if self.config_loader:
            raw_config = self.config_loader.get_raw_config()
            layout_bundle = raw_config.get('layout_bundle', {})
            for sheet_name, sheet_conf in layout_bundle.items():
                if not isinstance(sheet_conf, dict):
                    continue
                    
                # Check data_flow -> mappings -> col_desc
                mappings = sheet_conf.get('data_flow', {}).get('mappings', {})
                col_desc_rule = mappings.get('col_desc', {})
                if isinstance(col_desc_rule, dict):
                    fallback = col_desc_rule.get('fallback')
                    if isinstance(fallback, dict):
                        # Modern nested format
                        for mode_val in fallback.values():
                            if isinstance(mode_val, str) and mode_val.strip():
                                self.all_global_descriptions.add(mode_val.strip())
                    elif isinstance(fallback, str) and fallback.strip():
                        self.all_global_descriptions.add(fallback.strip())
        
        # Determine global uniqueness flag
        self.is_global_unique_desc = (len(self.all_global_descriptions) <= 1)
        # ---------------------------------------------

    @abstractmethod
    def process(self) -> bool:
        """
        Main method to orchestrate the processing of the worksheet.
        This must be implemented by all subclasses.

        Returns:
            bool: True if processing was successful, False otherwise.
        """
        pass
