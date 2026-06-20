from pydantic import BaseModel, Field
from typing import Dict, Optional, Any, List
from core.invoice_generator.styling.models import StylingConfigModel

class BorderExceptionsModel(BaseModel):
    default_border: str = "full_grid"
    default_style: str = "thin"
    exceptions: Dict[str, str] = Field(default_factory=dict)

class StylingDefaultsModel(BaseModel):
    borders: Optional[BorderExceptionsModel] = None

class CellStyleModel(BaseModel):
    format: Optional[str] = "@"
    alignment: Optional[str] = "center"
    wrap_text: bool = Field(False, alias="wrap_text")

class RowContextStyleModel(BaseModel):
    bold: bool = False
    font_size: int = 12
    font_name: str = "Arial"
    border_style: Optional[str] = "thin"
    add_ons: Optional[Dict[str, Any]] = None

class SheetStylingModel(BaseModel):
    columns: Dict[str, CellStyleModel] = Field(default_factory=dict)
    row_contexts: Dict[str, RowContextStyleModel] = Field(default_factory=dict)
