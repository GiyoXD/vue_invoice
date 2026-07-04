from abc import ABC, abstractmethod
from typing import List, Dict, Any, Set

from core.blueprint_generator.internal.scanner.models.addons import BaseAddonFact

class BaseAddonBuilder(ABC):
    """Abstract base class for all Addon Builders."""
    
    @abstractmethod
    def build_rows(self, fact: BaseAddonFact, sheet_col_ids: Set[str]) -> List[List[Dict[str, Any]]]:
        """
        Convert a pure fact into a list of JSON-formatted row dictionaries.
        
        Args:
            fact: The fact extracted by the scanner.
            sheet_col_ids: The set of column IDs present in the current sheet.
            
        Returns:
            A list of rows, where each row is a list of cell dictionaries.
        """
        pass
