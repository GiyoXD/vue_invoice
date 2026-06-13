from dataclasses import dataclass, field
from typing import List, Tuple, Optional

@dataclass
class TableZoneBoundary:
    """Represents the row boundaries for a single built table on the sheet."""
    table_key: str
    header_range: Tuple[int, int]
    data_range: Tuple[int, int]
    footer_range: Tuple[int, int]

@dataclass
class SheetLayoutState:
    """
    Tracks the layout of a sheet as it is built sequentially.
    Manages occupied row ranges to determine the next free row.
    """
    table_zones: List[TableZoneBoundary] = field(default_factory=list)
    
    # Template zones
    template_header_range: Optional[Tuple[int, int]] = None
    template_footer_range: Optional[Tuple[int, int]] = None
    
    # Internal tracker for the next available row on the sheet
    _next_free_row: int = 1
    
    @property
    def next_free_row(self) -> int:
        """Returns the next available row on the sheet that has not been built on."""
        return self._next_free_row
        
    def advance_to(self, row: int):
        """Advances the free row pointer."""
        self._next_free_row = max(self._next_free_row, row)

    def add_table_zone(self, table_key: str, header: Tuple[int, int], data: Tuple[int, int], footer: Tuple[int, int]):
        """Records a built table zone and advances the free row pointer."""
        self.table_zones.append(TableZoneBoundary(
            table_key=table_key,
            header_range=header,
            data_range=data,
            footer_range=footer
        ))
        
        # Advance the free row past the footer
        if footer and footer[1] >= footer[0]:
            self.advance_to(footer[1] + 1)
        elif data and data[1] >= data[0]:
            self.advance_to(data[1] + 1)
        elif header and header[1] >= header[0]:
            self.advance_to(header[1] + 1)
