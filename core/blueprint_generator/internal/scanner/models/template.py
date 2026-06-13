from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field

# Re-export classes moved to core.models.cell for backward compatibility
from core.models.cell import (
    FontStyle,
    AlignmentStyle,
    FillStyle,
    BorderStyle,
    CellStyle,
    TemplateMerge,
    UnitCell,
    UnitRow,
)

@dataclass
class TemplateLayout:
    header_rows: List[UnitRow] = field(default_factory=list)
    footer_rows: List[UnitRow] = field(default_factory=list)
    col_widths: Dict[str, float] = field(default_factory=dict)
    header_images: List[Dict[str, Any]] = field(default_factory=list)
    footer_images: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "header_rows": [r.to_dict() for r in self.header_rows],
            "footer_rows": [r.to_dict() for r in self.footer_rows],
            "col_widths": self.col_widths,
            "header_images": self.header_images,
            "footer_images": self.footer_images
        }

    @classmethod
    def from_dict(cls, d: Optional[Dict[str, Any]]) -> Optional['TemplateLayout']:
        if not d:
            return None
        return cls(
            header_rows=[UnitRow.from_dict(r) for r in d.get("header_rows", [])],
            footer_rows=[UnitRow.from_dict(r) for r in d.get("footer_rows", [])],
            col_widths=d.get("col_widths", {}),
            header_images=d.get("header_images", []),
            footer_images=d.get("footer_images", [])
        )
