import logging
from typing import Any, Dict, List, Optional, Union

from .summary_extractor import extract_summaries
from ..models import ResolvedTableFooter

logger = logging.getLogger(__name__)


class TableFooterMapper:
    """
    Mapper for resolving footer summaries and formatting display values.
    """
    
    def __init__(
        self,
        data_source_type: str,
        data_source: Union[Dict, List, None],
        footer_data: Dict[str, Any],
        table_key: Optional[str] = None
    ):
        self.data_source_type = data_source_type
        self.data_source = data_source
        self.footer_data = footer_data or {}
        self.table_key = table_key

    def resolve(self, data_rows: Optional[List[Dict[str, Any]]] = None, num_data_rows: int = 0) -> ResolvedTableFooter:
        """
        Resolves summary totals.
        """
        # Extract summaries if available in data source or footer data
        leather_summary, weight_summary, pallet_summary_total = extract_summaries(
            data_source=self.data_source,
            footer_data=self.footer_data,
            table_key=self.table_key
        )

        return ResolvedTableFooter(
            leather_summary=leather_summary,
            weight_summary=weight_summary,
            pallet_summary_total=pallet_summary_total
        )

    @staticmethod
    def create_from_bundles(
        data_config: Dict[str, Any],
        context_config: Dict[str, Any]
    ) -> 'TableFooterMapper':
        """
        Factory method to create TableFooterMapper from bundle configs.
        """
        return TableFooterMapper(
            data_source_type=data_config.get('data_source_type', 'aggregation'),
            data_source=data_config.get('data_source'),
            footer_data=data_config.get('footer_data', {}),
            table_key=data_config.get('table_key')
        )
