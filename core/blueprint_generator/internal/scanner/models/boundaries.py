from dataclasses import dataclass
from typing import Optional

@dataclass
class ZoneBoundaries:
    """Row boundaries for sheet zones."""
    header_row: int
    data_start_row: int
    footer_row: Optional[int] = None
    footer_end_row: Optional[int] = None
    max_col: int = 40

    @property
    def template_header_range(self) -> range:
        """Range of rows for template header content (Zone 1)."""
        return range(1, self.header_row)

    @property
    def table_range(self) -> range:
        """Range of rows for the entire table (Zone 2), from header to footer inclusive."""
        ref_row = self.footer_end_row if self.footer_end_row is not None else self.footer_row
        if ref_row is None:
            return range(self.header_row, self.header_row + 1)
        return range(self.header_row, ref_row + 1)

    def template_footer_range(self, max_row: int) -> range:
        """Range of rows for template footer content (Zone 3)."""
        ref_row = self.footer_end_row if self.footer_end_row is not None else self.footer_row
        if ref_row is None:
            return range(0, 0)
        limit_row = min(max_row, ref_row + 100)  # prevent ghost rows
        return range(ref_row + 1, limit_row + 1)
