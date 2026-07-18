import logging
from typing import Any, Dict, List, Tuple, Union, Optional

from .preparer import (
    prepare_data_rows,
    parse_mapping_rules
)
from .model import (
    ResolvedTableData,
    ResolvedTableFooter
)
from core.invoice_generator.models.config.layout import SheetLayoutModel
from .helpers import (
    extract_table_data,
    merge_static_content,
    extract_summaries,
    format_pallet_counts
)

logger = logging.getLogger(__name__)


class TableDataMapperError(Exception):
    """Exception raised when table data resolution fails."""
    pass


class TableDataMapper:
    """
    Mapper for preparing table-specific data for rendering.
    
    This class takes raw invoice data and configuration, then produces
    table-ready row dictionaries with proper formatting, formulas, and
    static values applied.
    """
    
    def __init__(
        self,
        data_source_type: str,
        data_source: Union[Dict, List, None],
        mapping_rules: Dict[str, Any],
        sheet_layout: Optional[SheetLayoutModel] = None,
        DAF_mode: bool = False,
        custom_mode: bool = False,
        static_content: Optional[Dict[str, Any]] = None,
        pricing_net_weight: bool = False,
        footer_data: Optional[Dict[str, Any]] = None,
        table_key: Optional[str] = None
    ):
        self.data_source_type = data_source_type
        self.data_source = data_source
        self.mapping_rules = mapping_rules
        self.sheet_layout = sheet_layout
        self.DAF_mode = DAF_mode
        self.custom_mode = custom_mode
        self.static_content = static_content or {}
        self.pricing_net_weight = pricing_net_weight
        self.footer_data = footer_data or {}
        self.table_key = table_key
        
        self.column_id_map = {}
        self.column_map = {}
        self.parent_column_ids = []
        
        if sheet_layout:
            bundled_columns, column_map, column_id_map, _ = (
                sheet_layout.structure.resolve_mappings(DAF_mode=DAF_mode, custom_mode=custom_mode)
            )
            self.column_id_map = column_id_map
            self.column_map = column_map
            self.parent_column_ids = [col.id for col in bundled_columns if col.children]
            
        self.idx_to_header_map = {v: k for k, v in self.column_map.items()}
        
        # Cached parsed rules
        self._parsed_rules = None
    
    def resolve(self) -> ResolvedTableData:
        """
        Main resolution method - transforms raw data into table-ready rows.
        
        Returns:
            ResolvedTableData model instance containing prepared rows.
        """
        # Parse mapping rules first
        parsed = self._parse_mapping_rules()
        
        # Extract data for this specific table (if multi-table)
        table_data_source = extract_table_data(self.data_source, self.data_source_type)
        
        # Prepare data rows using the existing data_preparer logic
        data_rows, num_data_rows = prepare_data_rows(
            data_source_type=self.data_source_type,
            data_source=table_data_source,
            dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
            column_id_map=self.column_id_map,
            idx_to_header_map=self.idx_to_header_map,
            desc_col_idx=self._get_desc_col_idx(),
            num_static_labels=parsed['num_static_labels'],
            static_value_map=parsed['static_value_map'],
            DAF_mode=self.DAF_mode,
            custom_mode=self.custom_mode,
            parent_column_ids=self.parent_column_ids,
            pricing_net_weight=self.pricing_net_weight
        )
        
        logger.debug(f"[DEBUG-RESOLVE] Parsed Rules: {parsed['dynamic_mapping_rules'].keys()}")
        logger.debug(f"[DEBUG-RESOLVE] column_id_map: {self.column_id_map}")
        logger.debug(f"[DEBUG-RESOLVE] idx_to_header: {self.idx_to_header_map}")
        logger.debug(f"[DEBUG-RESOLVE] Returned row 0 (if any): {data_rows[0] if data_rows else 'EMPTY'}")
        
        # Merge static content with data rows
        merge_static_content(
            data_rows=data_rows,
            static_content=self.static_content,
            dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
            DAF_mode=self.DAF_mode,
            custom_mode=self.custom_mode
        )
        
        # Resolve footer summaries internally
        footer_mapper = TableFooterMapper(
            data_source_type=self.data_source_type,
            data_source=self.data_source,
            footer_data=self.footer_data,
            table_key=self.table_key
        )
        resolved_footer = footer_mapper.resolve(data_rows, num_data_rows)
        
        return ResolvedTableData(
            data_rows=data_rows,
            num_data_rows=num_data_rows,
            footer=resolved_footer
        )
    
    def _parse_mapping_rules(self) -> Dict[str, Any]:
        """Parse mapping rules using existing data_preparer logic."""
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
        """
        Factory method to create TableDataMapper from bundle configs.
        """
        args = context_config.get('args')
        DAF_mode = args.DAF if args and hasattr(args, 'DAF') else False
        custom_mode = args.custom if args and hasattr(args, 'custom') else False
        
        static_content = {}
        if layout_config:
            static_content = layout_config.get('static_content', {})
            
        invoice_data = context_config.get('invoice_data') or {}
        metadata = invoice_data.get('metadata')
        
        if metadata is None:
            raise TableDataMapperError("CRITICAL: Invoice 'metadata' is missing or null in the provided JSON data.")
            
        pricing_net_weight = metadata.get('pricing_net_weight', False)
        
        sheet_layout = None
        if layout_config:
            sheet_layout = SheetLayoutModel.model_validate(layout_config.get('sheet_config', {}) or layout_config)

        return TableDataMapper(
            data_source_type=data_config.get('data_source_type', 'aggregation'),
            data_source=data_config.get('data_source'),
            mapping_rules=data_config.get('mapping_rules', {}),
            sheet_layout=sheet_layout,
            DAF_mode=DAF_mode,
            custom_mode=custom_mode,
            static_content=static_content,
            pricing_net_weight=pricing_net_weight,
            footer_data=data_config.get('footer_data', {}),
            table_key=data_config.get('table_key')
        )


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
