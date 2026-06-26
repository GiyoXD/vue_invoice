from pydantic import BaseModel, Field
from typing import Optional, List

class DataPrepHintModel(BaseModel):
    priority: List[str] = Field(default_factory=list)
    numbers_per_group_by_po: Optional[int] = None

class MetaModel(BaseModel):
    config_version: str
    customer: str
    created: Optional[str] = None
    description: Optional[str] = None
    has_static_sheets: bool = False
    data_preparation_module_hint: Optional[DataPrepHintModel] = None

class FeaturesModel(BaseModel):
    enable_text_replacement: bool = False
    enable_conditional_formatting: bool = False
    enable_data_validation: bool = False
    enable_auto_calculations: bool = True
    enable_print_area: bool = False
    enable_page_breaks: bool = False
    debug_mode: bool = False
