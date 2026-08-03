from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional

from .table import ColumnInfo, FooterInfo
from .template import TemplateLayout

@dataclass
class SheetAnalysis:
    """Complete analysis of a single sheet."""
    name: str
    header_row: int
    columns: List[ColumnInfo]
    data_source: str  # "aggregation" or "processed_tables_multi"
    header_font: Dict[str, Any]
    data_font: Dict[str, Any]
    row_heights: Dict[str, float]  # "header", "data", "footer" -> height
    has_multi_row_header: bool = False
    static_content_hints: Dict[str, Any] = field(default_factory=dict)
    static_layout: Optional[TemplateLayout] = None
    footer_info: Optional[FooterInfo] = None

    def to_legacy_dict(self) -> Dict[str, Any]:
        """Convert to legacy JSON format for frontend compatibility."""
        # Map columns to legacy header_positions
        header_positions = []
        for col in self.columns:
            header_positions.append({
                "keyword": col.header,
                "col_id": col.id,
                "row": self.header_row,
                "column": col.col_index
            })
            # Also add children if any
            if col.children:
                for child in col.children:
                    header_positions.append({
                        "keyword": child.header,
                        "row": self.header_row + 1,
                        "column": child.col_index
                    })

        return {
            "sheet_name": self.name,
            "header_positions": header_positions,
            "start_row": self.header_row + 1 if self.header_row else 1,
            "unconfirmed_footer": self.footer_info.total_text if (self.footer_info and not self.footer_info.is_exact) else None
        }


@dataclass
class TemplateAnalysisResult:
    """Complete template analysis result."""
    file_path: str
    customer_code: str
    sheets: List[SheetAnalysis]
    warnings: List[str] = field(default_factory=list)
    has_static_sheets: bool = False
    unrecognized_sheets: List[str] = field(default_factory=list)

    def to_legacy_dict(self) -> Dict[str, Any]:
        """Convert to legacy JSON format for frontend compatibility."""
        return {
            "file_path": self.file_path,
            "sheets": [sheet.to_legacy_dict() for sheet in self.sheets],
            "warnings": self.warnings,
            "has_static_sheets": self.has_static_sheets,
            "unrecognized_sheets": self.unrecognized_sheets
        }

