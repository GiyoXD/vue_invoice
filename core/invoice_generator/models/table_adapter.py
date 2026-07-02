from pydantic import BaseModel, Field
from typing import Dict, List, Optional, Any

class ResolvedTableData(BaseModel):
    data_rows: List[Dict[str, Any]] = Field(default_factory=list)
    num_data_rows: int = 0
    leather_summary: Optional[Dict[str, Any]] = None
    weight_summary: Optional[Dict[str, Any]] = None
    pallet_summary_total: Optional[int] = None
