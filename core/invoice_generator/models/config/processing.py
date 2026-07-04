from pydantic import RootModel
from typing import Dict, List

class ProcessingModel(RootModel[Dict[str, str]]):
    @property
    def sheets(self) -> List[str]:
        return list(self.root.keys())

    @property
    def data_sources(self) -> Dict[str, str]:
        return self.root
