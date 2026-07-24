import logging
from typing import Any, Dict, List, Optional, Union

from .models import ResolvedTableData, ResolvedTableFooter, MappingContext
from .binding import TableBinding
from .rules import parse_mapping_rules
from .transforms import (
    extract_table_data,
    prepare_data_rows,
    merge_static_content,
    extract_summaries,
    format_pallet_counts
)
from core.invoice_generator.models.config.layout import SheetLayoutModel

logger = logging.getLogger(__name__)


class TableDataMapperError(Exception):
    """Exception raised when table data resolution fails."""
    pass


class TableDataMapper:
    """
    Mapper for preparing table-specific data for rendering.
    Takes a unified TableBinding (or raw parameters) and produces table-ready rows.
    """
    
    def __init__(
        self,
        data_source_type: str = "aggregation",
        data_source: Union[Dict, List, None] = None,
        mapping_rules: Optional[Dict[str, Any]] = None,
        sheet_layout: Optional[SheetLayoutModel] = None,
        DAF_mode: bool = False,
        custom_mode: bool = False,
        static_content: Optional[Dict[str, Any]] = None,
        pricing_net_weight: bool = False,
        footer_data: Optional[Dict[str, Any]] = None,
        context: Optional[MappingContext] = None,
        binding: Optional[TableBinding] = None
    ):
        if binding is None:
            if context is not None:
                binding = TableBinding(
                    data=extract_table_data(context.data_source, context.data_source_type),
                    mapping_rules=context.mapping_rules,
                    sheet_layout=context.sheet_layout,
                    DAF_mode=context.DAF_mode,
                    custom_mode=context.custom_mode,
                    static_content=context.static_content,
                    pricing_net_weight=context.pricing_net_weight,
                    footer_data=context.footer_data,
                    data_source_type=context.data_source_type
                )
            else:
                binding = TableBinding(
                    data=extract_table_data(data_source, data_source_type),
                    mapping_rules=mapping_rules or {},
                    sheet_layout=sheet_layout,
                    DAF_mode=DAF_mode,
                    custom_mode=custom_mode,
                    static_content=static_content or {},
                    pricing_net_weight=pricing_net_weight,
                    footer_data=footer_data or {},
                    data_source_type=data_source_type
                )

        self.binding = binding
        self.context = context or MappingContext(
            data_source_type=binding.data_source_type,
            data_source=binding.data,
            mapping_rules=binding.mapping_rules,
            sheet_layout=binding.sheet_layout,
            DAF_mode=binding.DAF_mode,
            custom_mode=binding.custom_mode,
            static_content=binding.static_content,
            pricing_net_weight=binding.pricing_net_weight,
            footer_data=binding.footer_data
        )

        self.data_source_type = binding.data_source_type
        self.data_source = binding.data
        self.mapping_rules = binding.mapping_rules
        self.sheet_layout = binding.sheet_layout
        self.DAF_mode = binding.DAF_mode
        self.custom_mode = binding.custom_mode
        self.static_content = binding.static_content
        self.pricing_net_weight = binding.pricing_net_weight
        self.footer_data = binding.footer_data
        
        self.column_id_map = {}
        self.column_map = {}
        self.parent_column_ids = []
        
        if self.sheet_layout:
            bundled_columns, column_map, column_id_map, _ = (
                self.sheet_layout.structure.resolve_mappings(
                    DAF_mode=self.DAF_mode,
                    custom_mode=self.custom_mode
                )
            )
            self.column_id_map = column_id_map
            self.column_map = column_map
            self.parent_column_ids = [col.id for col in bundled_columns if col.children]
            
        self.idx_to_header_map = {v: k for k, v in self.column_map.items()}
        self._parsed_rules = None
    
    def resolve(self) -> ResolvedTableData:
        """Main resolution method - transforms raw data into clean ResolvedTableData."""
        return self.binding.resolve()

        
    def _parse_mapping_rules(self) -> Dict[str, Any]:
        """Parse mapping rules using existing logic."""
        if self._parsed_rules is None:
            self._parsed_rules = parse_mapping_rules(
                mapping_rules=self.mapping_rules,
                column_id_map=self.column_id_map,
                idx_to_header_map=self.idx_to_header_map
            )
        return self._parsed_rules
    
    def _get_desc_col_idx(self) -> int:
        """Get the description column index."""
        desc_col_id = None
        for possible_id in ['col_desc']:
            if possible_id in self.column_id_map:
                desc_col_id = possible_id
                break
        return self.column_id_map.get(desc_col_id, -1) if desc_col_id else -1
        
    @staticmethod
    def create_from_bundles(
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> 'TableDataMapper':
        """Factory method to create TableDataMapper from bundle configs."""
        args = context_config.get('args')
        DAF_mode = args.DAF if args and hasattr(args, 'DAF') else False
        custom_mode = args.custom if args and hasattr(args, 'custom') else False
        
        invoice_data = context_config.get('invoice_data') or {}
        metadata = invoice_data.get('metadata')
        if metadata is None:
            raise TableDataMapperError("CRITICAL: Invoice 'metadata' is missing or null in the provided JSON data.")
            
        binding = TableBinding.from_bundles(
            data_config=data_config,
            context_config={
                **context_config,
                'DAF_mode': DAF_mode,
                'custom_mode': custom_mode,
                'pricing_net_weight': metadata.get('pricing_net_weight', False)
            },
            layout_config=layout_config
        )

        return TableDataMapper(binding=binding)


class TableFooterMapper:
    """Mapper for resolving footer summaries and formatting display values."""
    
    def __init__(
        self,
        data_source_type: str = "aggregation",
        data_source: Union[Dict, List, None] = None,
        footer_data: Optional[Dict[str, Any]] = None,
        context: Optional[MappingContext] = None,
        binding: Optional[TableBinding] = None
    ):
        if binding is not None:
            self.binding = binding
            self.context = context or MappingContext(
                data_source_type=binding.data_source_type,
                data_source=binding.data,
                footer_data=binding.footer_data
            )
        elif context is not None:
            self.context = context
            self.binding = None
        else:
            self.context = MappingContext(
                data_source_type=data_source_type,
                data_source=data_source,
                footer_data=footer_data or {}
            )
            self.binding = None
        
        self.data_source_type = self.context.data_source_type
        self.data_source = self.context.data_source
        self.footer_data = self.context.footer_data

    def resolve(self, data_rows: Optional[List[Dict[str, Any]]] = None, num_data_rows: int = 0) -> ResolvedTableFooter:
        """Resolves summary totals."""
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
        """Factory method to create TableFooterMapper from bundle configs."""
        context = MappingContext.from_bundles(
            data_config=data_config,
            context_config=context_config
        )
        return TableFooterMapper(context=context)


def resolve_summary_payload(
    invoice_data: Optional[Dict[str, Any]] = None,
    footer_data_model: Optional[Any] = None
) -> Dict[str, Any]:
    """Resolves structured summary payload for summary sections."""
    invoice_data = invoice_data or {}
    footer_dict = invoice_data.get('footer_data', {})
    raw_grand_total = footer_dict.get('grand_total', {})

    pallet_count = raw_grand_total.get('col_pallet_count', raw_grand_total.get('pallet_count'))
    if pallet_count is None and footer_data_model:
        pallet_count = getattr(footer_data_model, 'total_pallets', 0)
    p_count = int(pallet_count or 0)

    grand_total_target = {
        **raw_grand_total,
        'pallet_count': p_count,
        'multiple': "S" if p_count != 1 else ""
    }

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

    payload = {
        'grand_total': grand_total_target,
        'weight_summary': weight_summary_target,
        'leather_summary': leather_summary_target
    }

    return payload
