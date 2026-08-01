from pydantic import RootModel
from typing import Dict, List, Optional, Union

class ProcessingModel(RootModel[Dict[str, Union[str, Dict[str, str]]]]):
    @property
    def sheets(self) -> List[str]:
        return list(self.root.keys())

    @property
    def data_sources(self) -> Dict[str, Union[str, Dict[str, str]]]:
        return self.root

    def get_source(self, sheet: str, mode: str = "standard") -> Optional[str]:
        """Resolve data source type for a sheet, optionally by mode.
        
        If the value is a dict (mode-aware), looks up by mode key with
        fallback to 'standard'. If it's a plain string, returns as-is.
        """
        val = self.root.get(sheet)
        if isinstance(val, dict):
            return val.get(mode, val.get("standard"))
        return val
