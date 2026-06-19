import logging
from abc import ABC, abstractmethod
from typing import Any

from .table_grid import Grid

logger = logging.getLogger(__name__)


class TableSectionBuilder(ABC):
    """
    Abstract base class for all table section builders (Header, Data, Footer).
    Workers now operate purely on a Grid, abstracting away physical coordinates.
    """
    def __init__(self, grid: Grid):
        self.grid = grid

    @abstractmethod
    def build(self) -> Any:
        """Build the section using the grid."""
        pass
