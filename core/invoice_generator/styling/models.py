from pydantic import BaseModel, Field
from typing import Optional, Dict, Union

class FontModel(BaseModel):
    name: Optional[str] = None
    size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color: Optional[str] = None

class AlignmentModel(BaseModel):
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    wrapText: Optional[bool] = False

class BorderStyleModel(BaseModel):
    style: Optional[str] = None
    color: Optional[str] = None

class ColumnStyleModel(BaseModel):
    font: Optional[FontModel] = None
    alignment: Optional[AlignmentModel] = None
    numberFormat: Optional[str] = None

class StylingConfigModel(BaseModel):
    defaultFont: Optional[FontModel] = Field(None, alias='default_font')
    defaultAlignment: Optional[AlignmentModel] = Field(None, alias='default_alignment')
    headerFont: Optional[FontModel] = Field(None, alias='header_font')
    headerAlignment: Optional[AlignmentModel] = Field(None, alias='header_alignment')
    columnIdStyles: Dict[str, ColumnStyleModel] = Field({}, alias='column_id_styles')
    columnIdsWithFullGrid: Optional[list[str]] = Field(None, alias='column_ids_with_full_grid')
    forceTextFormatIds: Optional[list[str]] = Field(None, alias='force_text_format_ids')
    columnIdWidths: Optional[Dict[str, float]] = Field(None, alias='column_id_widths')
    rowHeights: Optional[Dict[str, float]] = Field(None, alias='row_heights')

