from pydantic import BaseModel
from typing import Dict, List

class ProcessingModel(BaseModel):
    sheets: List[str]
    data_sources: Dict[str, str]
