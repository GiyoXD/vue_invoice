import logging
from typing import Any, Dict, List, Optional, Union

from .summary_extractor import extract_summaries
from ..models import ResolvedTableFooter, MappingContext

logger = logging.getLogger(__name__)


class TableFooterMapper:
    """
    Mapper for resolving footer summaries and formatting display values.
    """
    
    def __init__(
        self,
        data_source_type: str = "aggregation",
        data_source: Union[Dict, List, None] = None,
        footer_data: Optional[Dict[str, Any]] = None,
        context: Optional[MappingContext] = None
    ):
        if context is not None:
            self.context = context
        else:
            self.context = MappingContext(
                data_source_type=data_source_type,
                data_source=data_source,
                footer_data=footer_data or {}
            )
        
        self.data_source_type = self.context.data_source_type
        self.data_source = self.context.data_source
        self.footer_data = self.context.footer_data

    def resolve(self, data_rows: Optional[List[Dict[str, Any]]] = None, num_data_rows: int = 0) -> ResolvedTableFooter:
        """
        Resolves summary totals.
        """
        leather_summary, weight_summary, pallet_summary_total = extract_summaries(
            data_source=self.data_source,
            footer_data=self.footer_data
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
        context = MappingContext.from_bundles(
            data_config=data_config,
            context_config=context_config
        )
        return TableFooterMapper(context=context)


def resolve_summary_payload(
    invoice_data: Optional[Dict[str, Any]] = None,
    footer_data_model: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Resolves structured summary payload for summary sections:
      - grand_total: Grand total row metrics (col_pallet_count, col_amount, etc.)
      - weight_summary: Weight metrics (weight_net, weight_gross, col_net, col_gross)
      - leather_summary: Repeating leather breakdown records
    """
    invoice_data = invoice_data or {}
    footer_dict = invoice_data.get('footer_data', {})
    raw_grand_total = footer_dict.get('grand_total', {})

    # 1. Target: Grand Total metrics
    pallet_count = raw_grand_total.get('col_pallet_count', raw_grand_total.get('pallet_count'))
    if pallet_count is None and footer_data_model:
        pallet_count = getattr(footer_data_model, 'total_pallets', 0)
    p_count = int(pallet_count or 0)

    grand_total_target = {
        **raw_grand_total,
        'pallet_count': p_count,
        'multiple': "S" if p_count != 1 else ""
    }

    # 2. Target: Weight Summary metrics
    net_weight = float(raw_grand_total.get('col_net', 0.0))
    gross_weight = float(raw_grand_total.get('col_gross', 0.0))
    if not net_weight and footer_data_model and hasattr(footer_data_model, 'weight_summary'):
        ws = getattr(footer_data_model, 'weight_summary') or {}
        if isinstance(ws, dict):
            net_weight = float(ws.get('net', 0.0))
            gross_weight = float(ws.get('gross', 0.0))

    weight_summary_target = {
        'weight_net': net_weight,
        'weight_gross': gross_weight
    }

    # 3. Target: Leather Summary records (consistently enriched with pallet_count & multiple)
    raw_leather = footer_dict.get('leather_summary', [])
    leather_summary_target = []
    for item in raw_leather:
        if isinstance(item, dict):
            rec = dict(item)
            l_cnt = int(rec.get('col_pallet_count', rec.get('pallet_count', 0)) or 0)
            rec['pallet_count'] = l_cnt
            rec['multiple'] = "S" if l_cnt != 1 else ""
            leather_summary_target.append(rec)
        else:
            leather_summary_target.append(item)

    # Clean structured summary payload
    payload = {
        'grand_total': grand_total_target,
        'weight_summary': weight_summary_target,
        'leather_summary': leather_summary_target
    }

    return payload
