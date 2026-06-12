from .template import (
    TemplateLayout, UnitRow, UnitCell, TemplateMerge,
    CellStyle, FontStyle, AlignmentStyle, FillStyle, BorderStyle
)
from .boundaries import ZoneBoundaries
from .table import FooterInfo, ColumnInfo, TableLayout
from .analysis import SheetAnalysis, TemplateAnalysisResult

# Aliases for backward compatibility
TemplateRow = UnitRow
TemplateCell = UnitCell

__all__ = [
    "TemplateLayout",
    "UnitRow",
    "UnitCell",
    "TemplateRow",
    "TemplateCell",
    "TemplateMerge",
    "CellStyle",
    "FontStyle",
    "AlignmentStyle",
    "FillStyle",
    "BorderStyle",
    "ZoneBoundaries",
    "FooterInfo",
    "ColumnInfo",
    "TableLayout",
    "SheetAnalysis",
    "TemplateAnalysisResult",
]
