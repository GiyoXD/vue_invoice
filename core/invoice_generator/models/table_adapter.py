from pydantic import BaseModel, Field
from typing import Dict, List, Optional, Any

class StaticInfoModel(BaseModel):
    col1_index: Optional[int] = None
    num_static_labels: int = 0
    initial_static_col1_values: List[str] = Field(default_factory=list)
    static_column_header_name: Optional[str] = None
    apply_special_border_rule: bool = False

class ResolvedTableData(BaseModel):
    data_rows: List[Dict[str, Any]] = Field(default_factory=list)
    pallet_counts: List[int] = Field(default_factory=list)
    num_data_rows: int = 0
    static_info: StaticInfoModel = Field(default_factory=StaticInfoModel)
    formula_rules: Dict[str, Any] = Field(default_factory=dict)
    static_content: Dict[str, Any] = Field(default_factory=dict)
    leather_summary: Optional[Dict[str, Any]] = None
    weight_summary: Optional[Dict[str, Any]] = None
    pallet_summary_total: Optional[int] = None

    def __getitem__(self, item: str) -> Any:
        if item in self.__class__.model_fields:
            return getattr(self, item)
        raise KeyError(item)

    def get(self, item: str, default: Any = None) -> Any:
        if item in self.__class__.model_fields:
            return getattr(self, item)
        return default

    def __contains__(self, item: str) -> bool:
        return item in self.__class__.model_fields
