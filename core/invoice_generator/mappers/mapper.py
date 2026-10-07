import logging
from typing import Any, Dict, List, Optional, Union, Tuple

from .models import ResolvedTableData, ResolvedTableFooter, MappingContext, InvoiceMode
from .binding import TableBinding
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
        binding: Optional[TableBinding] = None,
    ):
        self.binding = binding or TableBinding()
        self.data_source = self.binding.data
        self.mapping_rules = self.binding.mapping_rules
        self.sheet_layout = self.binding.sheet_layout
        self.static_payload = self.binding.static_payload
        self.footer_data = self.binding.footer_data
        self.mode = self.binding.mode

        self.column_id_map = {}
        self.column_map = {}
        self.parent_column_ids = []
        
        if self.sheet_layout:
            bundled_columns, column_map, column_id_map, _ = (
                self.sheet_layout.structure.resolve_mappings(
                    DAF_mode=self.mode.DAF_mode,
                    custom_mode=self.mode.custom_mode
                )
            )
            self.column_id_map = column_id_map
            self.column_map = column_map
            self.parent_column_ids = [col.id for col in bundled_columns if col.children]
    
    def resolve(self) -> ResolvedTableData:
        """Main resolution method - transforms raw data into clean ResolvedTableData."""
        from .transforms import transform_table_data

        target_columns = [col for col in self.column_id_map if col not in self.parent_column_ids]
        data_rows, num_data_rows = transform_table_data(
            data_source=self.data_source,
            mapping_rules=self.mapping_rules,
            static_payload=self.static_payload,
            mode=self.mode,
            target_columns=target_columns
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



    @staticmethod
    def create_from_bundles(
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> 'TableDataMapper':
        """Factory method to create TableDataMapper from bundle configs via TableBinding gate."""
        binding = TableBinding.from_bundles(
            data_config=data_config,
            context_config=context_config,
            layout_config=layout_config
        )
        return TableDataMapper(binding=binding)

def extract_summaries(
    data_source: Any = None,
    footer_data: Optional[Dict[str, Any]] = None,
    table_key: Optional[str] = None
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], Optional[int]]:
    """Extract leather_summary, weight_summary, and pallet_summary_total from footer_data."""
    footer = footer_data or {}
    add_ons = footer.get('add_ons') or {}
    leather_summary = footer.get('leather_summary') or add_ons.get('leather_summary_addon')
    weight_summary = footer.get('weight_summary') or add_ons.get('weight_summary_addon')
    pallet_summary_total = None
            
    grand_total = footer.get('grand_total') or {}
    if 'col_pallet_count' in grand_total and table_key is None:
        pallet_summary_total = grand_total.get('col_pallet_count')
    
    if pallet_summary_total is None:
        table_totals = footer.get('table_totals')
        if table_totals is not None:
            if not isinstance(table_totals, list):
                raise TypeError(f"Invalid 'table_totals' type: expected list of dicts, got {type(table_totals).__name__}")
            idx = 0
            if table_key is not None:
                try:
                    idx = int(table_key)
                except (ValueError, TypeError):
                    idx = 0
            if 0 <= idx < len(table_totals):
                tbl_footer = table_totals[idx] or {}
                pallet_summary_total = tbl_footer.get('col_pallet_count')
            
    return leather_summary, weight_summary, pallet_summary_total


class TableFooterMapper:
    """Mapper for resolving footer summaries and formatting display values."""
    
    def __init__(
        self,
        binding: Optional[TableBinding] = None,
    ):
        self.binding = binding or TableBinding()
        self.data_source = self.binding.data
        self.footer_data = self.binding.footer_data
        self.mode = self.binding.mode
        self.context = MappingContext(
            data_source=self.data_source,
            footer_data=self.footer_data,
            invoice_format=self.mode.invoice_format,
            pricing_basis=self.mode.pricing_basis
        )

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
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> 'TableFooterMapper':
        """Factory method to create TableFooterMapper from bundle configs via TableBinding gate."""
        binding = TableBinding.from_bundles(
            data_config=data_config,
            context_config=context_config,
            layout_config=layout_config
        )
        return TableFooterMapper(binding=binding)


def resolve_flat_footer_payload(
    raw_dict: Optional[Dict[str, Any]] = None,
    footer_data_model: Optional[Any] = None
) -> Dict[str, Any]:
    """Passes flat payload directly from Data Parser."""
    raw = dict(raw_dict or {})
    if raw.get('col_pallet_count') is None and footer_data_model:
        raw['col_pallet_count'] = getattr(footer_data_model, 'total_pallets', 0)
    raw_pallets = raw.get('col_pallet_count')
    pallet_cnt = int(raw_pallets) if raw_pallets is not None else 0
    raw.setdefault('multiple', "S" if pallet_cnt != 1 else "")
    return raw


def resolve_summary_payload(
    invoice_data: Optional[Dict[str, Any]] = None,
    footer_data_model: Optional[Any] = None
) -> Dict[str, Any]:
    """Resolves structured summary payload for summary sections."""
    invoice_data = invoice_data or {}
    footer_dict = invoice_data.get('footer_data') or {}
    raw_grand_total = footer_dict.get('grand_total') or {}

    grand_total_target = resolve_flat_footer_payload(raw_grand_total, footer_data_model)

    net_val = raw_grand_total.get('col_net')
    gross_val = raw_grand_total.get('col_gross')
    net_weight = float(net_val) if net_val is not None else 0.0
    gross_weight = float(gross_val) if gross_val is not None else 0.0
    if not net_weight and footer_data_model and hasattr(footer_data_model, 'weight_summary'):
        ws = getattr(footer_data_model, 'weight_summary') or {}
        ws_net = ws.get('net')
        ws_gross = ws.get('gross')
        net_weight = float(ws_net) if ws_net is not None else 0.0
        gross_weight = float(ws_gross) if ws_gross is not None else 0.0

    weight_summary_target = {
        'weight_net': net_weight,
        'weight_gross': gross_weight
    }

    leather_summary_target = footer_dict.get('leather_summary') or []

    return {
        'grand_total': grand_total_target,
        'weight_summary': weight_summary_target,
        'leather_summary': leather_summary_target
    }


__all__ = [
    "TableDataMapperError",
    "TableDataMapper",
    "TableFooterMapper",
    "extract_summaries",
    "resolve_flat_footer_payload",
    "resolve_summary_payload",
]


