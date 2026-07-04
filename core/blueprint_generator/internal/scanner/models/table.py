from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional

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
class TableLayout:
    """Complete table zone analysis (Zone 2): structure, styling, and footer."""
    header_row: int
    data_start_row: int
    columns: List[ColumnInfo]
    header_font: Dict[str, Any]
    data_font: Dict[str, Any]
    row_heights: Dict[str, float]
    has_multi_row_header: bool
    footer_info: Optional[FooterInfo] = None
    static_content_hints: Dict[str, List[str]] = field(default_factory=dict)
