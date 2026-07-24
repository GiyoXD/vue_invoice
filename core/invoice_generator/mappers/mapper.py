import logging
from typing import Any, Dict, List, Optional, Union

from .models import ResolvedTableData, ResolvedTableFooter, MappingContext
from .binding import TableBinding
from .rules import parse_mapping_rules
from .transforms import (
    extract_table_data,
    prepare_data_rows,
    populate_static_content,
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
                options = {
                    'DAF_mode': context.DAF_mode,
                    'custom_mode': context.custom_mode,
                    'pricing_net_weight': context.pricing_net_weight
                }
                binding = TableBinding(
                    data=extract_table_data(context.data_source, context.data_source_type),
                    mapping_rules=context.mapping_rules,
                    sheet_layout=context.sheet_layout,
                    static_payload=context.static_payload,
                    footer_data=context.footer_data,
                    data_source_type=context.data_source_type,
                    options=options
                )
            else:
                options = {
                    'DAF_mode': DAF_mode,
                    'custom_mode': custom_mode,
                    'pricing_net_weight': pricing_net_weight
                }
                binding = TableBinding(
                    data=extract_table_data(data_source, data_source_type),
                    mapping_rules=mapping_rules or {},
                    sheet_layout=sheet_layout,
                    static_payload=static_content or {},
                    footer_data=footer_data or {},
                    data_source_type=data_source_type,
                    options=options
                )

        self.binding = binding
        self.context = context or MappingContext(
            data_source_type=binding.data_source_type,
            data_source=binding.data,
            mapping_rules=binding.mapping_rules,
            sheet_layout=binding.sheet_layout,
            DAF_mode=binding.options.get('DAF_mode', False),
            custom_mode=binding.options.get('custom_mode', False),
            static_payload=binding.static_payload,
            pricing_net_weight=binding.options.get('pricing_net_weight', False),
            footer_data=binding.footer_data
        )

        self.data_source_type = binding.data_source_type
        self.data_source = binding.data
        self.mapping_rules = binding.mapping_rules
        self.sheet_layout = binding.sheet_layout
        self.static_payload = binding.static_payload
        self.footer_data = binding.footer_data
        
        self.column_id_map = {}
        self.column_map = {}
        self.parent_column_ids = []
        
        if self.sheet_layout:
            bundled_columns, column_map, column_id_map, _ = (
                self.sheet_layout.structure.resolve_mappings(
                    DAF_mode=self.binding.options.get('DAF_mode', False),
                    custom_mode=self.binding.options.get('custom_mode', False)
                )
            )
            self.column_id_map = column_id_map
            self.column_map = column_map
            self.parent_column_ids = [col.id for col in bundled_columns if col.children]
            
        self.idx_to_header_map = {v: k for k, v in self.column_map.items()}
        self._parsed_rules = None
    
    def resolve(self) -> ResolvedTableData:
        """Main resolution method - transforms raw data into clean ResolvedTableData."""
        from .rules import parse_mapping_rules, get_fallback_value
        from .transforms import (
            prepare_data_rows,
            resolve_static_placeholders,
            populate_static_content,
            format_pallet_counts,
            extract_summaries
        )

        DAF_mode = self.binding.options.get('DAF_mode', False)
        custom_mode = self.binding.options.get('custom_mode', False)
        pricing_net_weight = self.binding.options.get('pricing_net_weight', False)

        column_id_map = {}
        column_map = {}
        parent_column_ids = []

        if self.sheet_layout:
            bundled_columns, column_map, column_id_map, _ = (
                self.sheet_layout.structure.resolve_mappings(
                    DAF_mode=DAF_mode,
                    custom_mode=custom_mode
                )
            )
            parent_column_ids = [col.id for col in bundled_columns if col.children]

        idx_to_header_map = {v: k for k, v in column_map.items()}

        parsed = parse_mapping_rules(
            mapping_rules=self.mapping_rules,
            column_id_map=column_id_map,
            idx_to_header_map=idx_to_header_map
        )

        # 1. Resolve description fallback string
        desc_rule = parsed['dynamic_mapping_rules'].get('col_desc', {})
        desc_fallback_str = get_fallback_value(desc_rule, DAF_mode, custom_mode) or ""

        # 2. Replace placeholders in static payload
        resolved_static_payload = resolve_static_placeholders(
            static_payload=self.static_payload,
            desc_fallback_str=str(desc_fallback_str)
        )

        # 3. Prepare data rows
        desc_col_id = 'col_desc' if 'col_desc' in column_id_map else None
        desc_col_idx = column_id_map.get(desc_col_id, -1) if desc_col_id else -1

        data_rows, num_data_rows = prepare_data_rows(
            data_source_type=self.data_source_type,
            data_source=self.data_source,
            dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
            column_id_map=column_id_map,
            idx_to_header_map=idx_to_header_map,
            desc_col_idx=desc_col_idx,
            num_static_labels=parsed['num_static_labels'],
            static_value_map=parsed['static_value_map'],
            DAF_mode=DAF_mode,
            custom_mode=custom_mode,
            parent_column_ids=parent_column_ids,
            pricing_net_weight=pricing_net_weight
        )

        # 4. Populate resolved static payload into data rows
        populate_static_content(
            data_rows=data_rows,
            static_payload=resolved_static_payload
        )

        pallet_col_id = 'col_pallet_count' if 'col_pallet_count' in column_id_map else None
        if pallet_col_id:
            format_pallet_counts(
                data_rows=data_rows,
                num_data_rows=num_data_rows,
                pallet_col_id=pallet_col_id
            )

        leather_summary, weight_summary, pallet_summary_total = extract_summaries(
            data_source=self.data_source,
            footer_data=self.footer_data
        )

        resolved_footer = ResolvedTableFooter(
            leather_summary=leather_summary,
            weight_summary=weight_summary,
            pallet_summary_total=pallet_summary_total
        )

        return ResolvedTableData(
            data_rows=data_rows,
            num_data_rows=num_data_rows,
            footer=resolved_footer
        )

        
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
