import logging
from typing import Optional, Dict, Any
from pydantic import BaseModel, Field, ConfigDict

logger = logging.getLogger(__name__)


class FooterData(BaseModel):
    """
    Data object passed from DataTableBuilder to TableFooterBuilder.
    Contains all necessary information to render the footer without further calculation.
    """
    footer_row_start_idx: int
    data_start_row: int
    data_end_row: int
    total_pallets: int
    leather_summary: Optional[Any] = None
    weight_summary: Optional[Any] = None
