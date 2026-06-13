from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field

def _filter_empty(d: Dict[str, Any]) -> Dict[str, Any]:
    """Remove keys with values of None, False, or empty string."""
    return {k: v for k, v in d.items() if v not in (None, False, "")}

# ==============================================================================
# STYLE SUB-MODELS (Visual Styling of Cells)
# ==============================================================================

@dataclass
class FontStyle:
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return _filter_empty({
            "name": self.name,
            "size": self.size,
            "bold": self.bold,
            "italic": self.italic,
            "color": self.color,
        })

    @classmethod
    def from_dict(cls, d: Optional[Dict[str, Any]]) -> Optional['FontStyle']:
        if not d:
            return None
        return cls(
            name=d.get("name"),
            size=d.get("size"),
            bold=d.get("bold", False),
            italic=d.get("italic", False),
            color=d.get("color")
        )

@dataclass
class AlignmentStyle:
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    wrap_text: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return _filter_empty({
            "horizontal": self.horizontal,
            "vertical": self.vertical,
            "wrap_text": self.wrap_text,
        })

    @classmethod
    def from_dict(cls, d: Optional[Dict[str, Any]]) -> Optional['AlignmentStyle']:
        if not d:
            return None
        return cls(
            horizontal=d.get("horizontal"),
            vertical=d.get("vertical"),
            wrap_text=d.get("wrap_text", False)
        )

@dataclass
class FillStyle:
    fill_type: Optional[str] = None
    color: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return _filter_empty({
            "type": self.fill_type,
            "color": self.color,
        })

    @classmethod
    def from_dict(cls, d: Optional[Dict[str, Any]]) -> Optional['FillStyle']:
        if not d:
            return None
        return cls(
            fill_type=d.get("type"),
            color=d.get("color")
        )

@dataclass
class BorderStyle:
    left: Optional[str] = None
    right: Optional[str] = None
    top: Optional[str] = None
    bottom: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return _filter_empty({
            "left": self.left,
            "right": self.right,
            "top": self.top,
            "bottom": self.bottom,
        })

    @classmethod
    def from_dict(cls, d: Optional[Dict[str, Any]]) -> Optional['BorderStyle']:
        if not d:
            return None
        return cls(
            left=d.get("left"),
            right=d.get("right"),
            top=d.get("top"),
            bottom=d.get("bottom")
        )

@dataclass
class CellStyle:
    font: Optional[FontStyle] = None
    alignment: Optional[AlignmentStyle] = None
    fill: Optional[FillStyle] = None
    border: Optional[BorderStyle] = None
    number_format: str = "General"

    def to_dict(self) -> Dict[str, Any]:
        return _filter_empty({
            "font": self.font.to_dict() if self.font else None,
            "alignment": self.alignment.to_dict() if self.alignment else None,
            "fill": self.fill.to_dict() if self.fill else None,
            "border": self.border.to_dict() if self.border else None,
            "number_format": self.number_format if self.number_format != "General" else None,
        })

    @classmethod
    def from_dict(cls, d: Optional[Dict[str, Any]]) -> Optional['CellStyle']:
        if not d:
            return None
        return cls(
            font=FontStyle.from_dict(d.get("font")),
            alignment=AlignmentStyle.from_dict(d.get("alignment")),
            fill=FillStyle.from_dict(d.get("fill")),
            border=BorderStyle.from_dict(d.get("border")),
            number_format=d.get("number_format", "General")
        )

# ==============================================================================
# LAYOUT STRUCTURE MODELS (Positioning, Merges, Rows, Grid Layout)
# ==============================================================================

@dataclass
class TemplateMerge:
    min_col: int
    max_col: int
    row_span: int = 1
    value: str = ""

@dataclass
class UnitCell:
    col_index: int
    value: Optional[Any] = None
    style: Optional[CellStyle] = None
    merge: Optional[TemplateMerge] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            k: v for k, v in {
                "col_index": self.col_index,
                "value": self.value,
                "style": self.style.to_dict() if self.style else None,
                "merge": {
                    "min_col": self.merge.min_col,
                    "max_col": self.merge.max_col,
                    "row_span": self.merge.row_span,
                    "value": self.merge.value
                } if self.merge else None
            }.items() if v is not None
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'UnitCell':
        merge_data = d.get("merge")
        return cls(
            col_index=d["col_index"],
            value=d.get("value"),
            style=CellStyle.from_dict(d.get("style")) if "style" in d else None,
            merge=TemplateMerge(**merge_data) if merge_data else None
        )

@dataclass
class UnitRow:
    relative_index: int
    height: Optional[float] = None
    cells: List[UnitCell] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "relative_index": self.relative_index,
            "height": self.height,
            "cells": [c.to_dict() for c in self.cells]
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'UnitRow':
        return cls(
            relative_index=d["relative_index"],
            height=d.get("height"),
            cells=[UnitCell.from_dict(c) for c in d.get("cells", [])]
        )
