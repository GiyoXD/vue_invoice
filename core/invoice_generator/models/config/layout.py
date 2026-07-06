from pydantic import BaseModel, Field
from typing import Dict, List, Optional, Any, Union

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

    def resolve_mappings(self, DAF_mode: bool = False, custom_mode: bool = False):
        """
        Resolves active columns and computes mappings based on mode filters.
        """
        column_index_mapping = {}
        bundled_columns = []
        
        if not self.columns:
            return [], {}, {}, {}
            
        template_col = 1
        output_col = 1
        
        for col_def in self.columns:
            skip_daf = col_def.skip_in_daf
            skip_custom = col_def.skip_in_custom
            colspan_val = col_def.colspan
            children_list = col_def.children
            
            num_columns = len(children_list) if children_list else colspan_val
            should_skip = (DAF_mode and skip_daf) or (custom_mode and skip_custom)
            
            if should_skip:
                for i in range(num_columns):
                    column_index_mapping[template_col + i] = None
            else:
                for i in range(num_columns):
                    column_index_mapping[template_col + i] = output_col + i
                output_col += num_columns
                bundled_columns.append(col_def)
            
            template_col += num_columns

        # Convert logical ID to physical column mapping using filtered column layout
        col_id_mapping = {}
        col_colspan = {}
        col_index = 1
        for col in bundled_columns:
            col_id = col.id
            if col.children:
                # Parent columns should not be horizontally merged in data rows
                col_colspan[col_id] = 1
                col_id_mapping[col_id] = col_index
                for child in col.children:
                    child_id = child.id
                    col_id_mapping[child_id] = col_index
                    col_colspan[child_id] = 1
                    col_index += 1
            else:
                colspan = col.colspan
                col_id_mapping[col_id] = col_index
                col_colspan[col_id] = colspan
                col_index += colspan

        return bundled_columns, column_index_mapping, col_id_mapping, col_colspan

class MappingRuleModel(BaseModel):
    column: Optional[str] = None
    fallback_on_none: Optional[str] = None
    fallback_on_DAF: Optional[str] = None
    source_value: Optional[str] = None

class DataFlowConfigModel(BaseModel):
    mappings: Dict[str, MappingRuleModel] = Field(default_factory=dict)

class StaticContentConfigModel(BaseModel):
    static: Dict[str, List[str]] = Field(default_factory=dict)

class FooterConfigModel(BaseModel):
    type: str = "regular"
    rows: List[Union[List[Dict[str, Any]], Dict[str, Any]]] = Field(default_factory=list)

class HSCodeConfigModel(BaseModel):
    col_id: str
    value: str
    colspan: int = 1
    style_context: str = "footer"

class SheetLayoutModel(BaseModel):
    structure: StructureConfigModel
    data_flow: DataFlowConfigModel = Field(default_factory=DataFlowConfigModel)
    content: Optional[StaticContentConfigModel] = None
    footer: Optional[FooterConfigModel] = None
    hs_code: Optional[HSCodeConfigModel] = None
