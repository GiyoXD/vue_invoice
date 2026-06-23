from pydantic import BaseModel, Field
from typing import Dict, List, Optional, Any

class ColumnDef(BaseModel):
    id: str
    header: str = ""
    width: Optional[float] = None
    rowspan: int = 1
    colspan: int = 1
    source_field: Optional[str] = None
    skip_in_daf: bool = False
    skip_in_custom: bool = False
    children: List["ColumnDef"] = Field(default_factory=list)

class StructureConfigModel(BaseModel):
    header_row: int = 1
    columns: List[ColumnDef] = Field(default_factory=list)

class MappingRuleModel(BaseModel):
    column: str
    fallback_on_none: Optional[str] = None
    fallback_on_DAF: Optional[str] = None
    source_value: Optional[str] = None

class DataFlowConfigModel(BaseModel):
    mappings: Dict[str, MappingRuleModel] = Field(default_factory=dict)

class StaticContentConfigModel(BaseModel):
    static: Dict[str, List[str]] = Field(default_factory=dict)

class FooterMergeRuleModel(BaseModel):
    start_column_id: str
    colspan: int
    comment: Optional[str] = None

class FooterConfigModel(BaseModel):
    total_text_column_id: Optional[str] = None
    total_text: str = "TOTAL:"
    pallet_count_column_id: Optional[str] = None
    sum_column_ids: List[str] = Field(default_factory=list)
    sum_cols: List[str] = Field(default_factory=list)
    footer_cells: List[List[Any]] = Field(default_factory=list)
    add_blank_before: bool = False
    type: str = "regular"
    merge_rules: List[FooterMergeRuleModel] = Field(default_factory=list)
    add_ons: Optional[Dict[str, Any]] = None

class SheetLayoutModel(BaseModel):
    structure: StructureConfigModel
    data_flow: DataFlowConfigModel = Field(default_factory=DataFlowConfigModel)
    content: Optional[StaticContentConfigModel] = None
    footer: Optional[FooterConfigModel] = None
