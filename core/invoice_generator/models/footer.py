import logging
from typing import Optional, Dict
from pydantic import BaseModel, Field, ConfigDict

logger = logging.getLogger(__name__)


class WeightDetail(BaseModel):
    net: Optional[float] = 0.0
    gross: Optional[float] = 0.0

    model_config = ConfigDict(populate_by_name=True, extra='allow')


class LeatherDetail(BaseModel):
    col_qty_pcs: Optional[int] = Field(0, alias='pcs')
    col_qty_sf: Optional[float] = Field(0.0, alias='sqft')
    col_net: Optional[float] = Field(0.0, alias='net')
    col_gross: Optional[float] = Field(0.0, alias='gross')
    col_cbm: Optional[float] = Field(0.0, alias='cbm')
    col_pallet_count: Optional[int] = Field(0, alias='pallet_count')

    model_config = ConfigDict(populate_by_name=True, extra='allow')


class FooterData(BaseModel):
    """
    Data object passed from DataTableBuilder to TableFooterBuilder.
    Contains all necessary information to render the footer without further calculation.
    """
    footer_row_start_idx: int
    data_start_row: int
    data_end_row: int
    total_pallets: int
    leather_summary: Optional[Dict[str, LeatherDetail]] = None
    weight_summary: Optional[WeightDetail] = None
