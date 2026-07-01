from dataclasses import dataclass, field
from abc import ABC

@dataclass
class BaseAddonFact(ABC):
    """Abstract base class for all addon facts discovered by the scanner."""
    fact_type: str = field(init=False)

@dataclass
class LeatherSummaryFact(BaseAddonFact):
    """Pure facts about a detected leather summary row."""
    leather_key: str  # "BUFFALO" or "COW"
    total_col_id: str
    label_col_id: str
    total_value: str
    label_value: str
    
    def __post_init__(self) -> None:
        self.fact_type = "leather_summary"

