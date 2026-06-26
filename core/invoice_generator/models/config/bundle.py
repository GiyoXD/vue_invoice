from pydantic import BaseModel, Field, ConfigDict, model_validator
from typing import Dict, Optional, Union, Any
from .meta import MetaModel, FeaturesModel
from .processing import ProcessingModel
from .styling import StylingDefaultsModel, SheetStylingModel
from .layout import SheetLayoutModel, FooterConfigModel

class GlobalDefaultsModel(BaseModel):
    footer: Optional[FooterConfigModel] = None

class ClientConfigBundle(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    meta: MetaModel = Field(..., alias="_meta")
    features: Optional[FeaturesModel] = None
    processing: ProcessingModel
    styling_bundle: Dict[str, Union[StylingDefaultsModel, SheetStylingModel]]
    layout_bundle: Dict[str, Union[GlobalDefaultsModel, SheetLayoutModel]]
    defaults: Optional[GlobalDefaultsModel] = None

    @model_validator(mode='before')
    @classmethod
    def clean_comments(cls, data: Any) -> Any:
        if isinstance(data, dict):
            # Clean styling_bundle comments
            sb = data.get("styling_bundle")
            if isinstance(sb, dict):
                data["styling_bundle"] = {k: v for k, v in sb.items() if not k.startswith("_comment")}
            # Clean layout_bundle comments
            lb = data.get("layout_bundle")
            if isinstance(lb, dict):
                data["layout_bundle"] = {k: v for k, v in lb.items() if not k.startswith("_comment")}
        return data

