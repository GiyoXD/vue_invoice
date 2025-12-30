# invoice_generator/processors/base_processor.py
from abc import ABC, abstractmethod
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import argparse
from typing import Dict, Any, Optional

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

    @abstractmethod
    def process(self) -> bool:
        """
        Main method to orchestrate the processing of the worksheet.
        This must be implemented by all subclasses.

        Returns:
            bool: True if processing was successful, False otherwise.
        """
        pass
