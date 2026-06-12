from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field

@dataclass
class FooterInfo:
    """Information about the footer structure."""
    row_num: int
    total_text: str
    total_text_col_id: str
    merge_curr_colspan: int
    pallet_count_col_id: Optional[str] = None
    has_hs_code: bool = False
    hs_code_text: Optional[str] = None
    hs_code_colspan: int = 1
    hs_code_col_id: Optional[str] = None
    is_exact: bool = True


@dataclass
class ZoneBoundaries:
    """Row boundaries for sheet zones."""
    header_row: int
    data_start_row: int
    footer_row: Optional[int] = None
    max_col: int = 40

    @property
    def template_header_range(self) -> range:
        """Range of rows for template header content (Zone 1)."""
        return range(1, self.header_row)

    @property
    def table_range(self) -> range:
        """Range of rows for the entire table (Zone 2), from header to footer inclusive."""
        if self.footer_row is None:
            return range(self.header_row, self.header_row + 1)
        return range(self.header_row, self.footer_row + 1)

    def template_footer_range(self, max_row: int) -> range:
        """Range of rows for template footer content (Zone 3)."""
        if self.footer_row is None:
            return range(0, 0)
        limit_row = min(max_row, self.footer_row + 100)  # prevent ghost rows
        return range(self.footer_row + 1, limit_row + 1)



@dataclass
class TableLayout:
    """Complete table zone analysis (Zone 2): structure, styling, and footer."""
    header_row: int
    data_start_row: int
    columns: List['ColumnInfo']
    header_font: Dict[str, Any]
    data_font: Dict[str, Any]
    row_heights: Dict[str, float]
    has_multi_row_header: bool
    footer_info: Optional[FooterInfo] = None
    static_content_hints: Dict[str, List[str]] = field(default_factory=dict)


@dataclass
class ColumnInfo:
    """Information about a single column."""
    id: str
    header: str
    col_index: int  # 1-based
    width: float
    format: str = "@"
    alignment: str = "center"
    rowspan: int = 1
    colspan: int = 1
    children: List['ColumnInfo'] = field(default_factory=list)
    wrap_text: bool = False


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
    static_content_hints: Dict[str, List[str]] = field(default_factory=dict)
    static_layout: Optional[Dict[str, Any]] = None
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

    def to_legacy_dict(self) -> Dict[str, Any]:
        """Convert to legacy JSON format for frontend compatibility."""
        return {
            "file_path": self.file_path,
            "sheets": [sheet.to_legacy_dict() for sheet in self.sheets],
            "warnings": self.warnings,
            "has_static_sheets": self.has_static_sheets
        }
