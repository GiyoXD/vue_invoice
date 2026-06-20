from pydantic import BaseModel, Field
from typing import Dict, List, Optional, Any

class ColumnDef(BaseModel):
    id: str
    header: str
    width: Optional[float] = None
    rowspan: int = 1
    colspan: int = 1
    format: Optional[str] = None
    children: List["ColumnDef"] = Field(default_factory=list)

class StructureConfigModel(BaseModel):
    header_row: int
    row_heights: Dict[str, float] = Field(default_factory=dict)
    columns: List[ColumnDef] = Field(default_factory=list)

class MappingRuleModel(BaseModel):
    column: str
    fallback_on_none: Optional[str] = Field(None, alias="fallback_on_none")
    fallback_on_DAF: Optional[str] = Field(None, alias="fallback_on_DAF")
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
    total_text_column_id: str
    total_text: str = "TOTAL:"
    pallet_count_column_id: Optional[str] = None
    sum_column_ids: List[str] = Field(default_factory=list)
    merge_rules: List[FooterMergeRuleModel] = Field(default_factory=list)
    add_ons: Optional[Dict[str, Any]] = None

class SheetLayoutModel(BaseModel):
    sections: List[str] = Field(default_factory=list, alias="_sections")
    structure: StructureConfigModel
    data_flow: DataFlowConfigModel
    content: Optional[StaticContentConfigModel] = None
    footer: Optional[FooterConfigModel] = None
