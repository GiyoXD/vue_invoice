from dataclasses import dataclass
from typing import Dict, Any, Optional
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

@dataclass
class ExcelIOContext:
    """Holds references to Excel objects (reading template, writing output)."""
    template_workbook: Workbook
    output_workbook: Workbook
    template_worksheet: Worksheet
    output_worksheet: Worksheet

@dataclass
class SheetConfigContext:
    """Holds all configuration and routing logic for the sheet."""
    sheet_name: str
    sheet_config: Dict[str, Any]
    data_source_indicator: str
    config_loader: Optional[Any] = None

@dataclass
class RuntimeDataContext:
    """Holds the raw input data and runtime arguments."""
    invoice_data: Dict[str, Any]
    cli_args: Any
    final_grand_total_pallets: int

@dataclass
class ProcessorContext:
    """Aggregates the domains for the Processor."""
    io: ExcelIOContext
    config: SheetConfigContext
    data: RuntimeDataContext

