import logging
from typing import Any, Dict, List, Optional, Tuple, Union

from .row_builder import prepare_data_rows
from .static_merger import merge_static_content
from ..rules import parse_mapping_rules
from ..transforms import extract_table_data
from ..models import ResolvedTableData, MappingContext
from core.invoice_generator.models.config.layout import SheetLayoutModel

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
        data_source_type: str = "aggregation",
        data_source: Union[Dict, List, None] = None,
        mapping_rules: Optional[Dict[str, Any]] = None,
        sheet_layout: Optional[SheetLayoutModel] = None,
        DAF_mode: bool = False,
        custom_mode: bool = False,
        static_content: Optional[Dict[str, Any]] = None,
        pricing_net_weight: bool = False,
        footer_data: Optional[Dict[str, Any]] = None,
        context: Optional[MappingContext] = None
    ):
        if context is not None:
            self.context = context
        else:
            self.context = MappingContext(
                data_source_type=data_source_type,
                data_source=data_source,
                mapping_rules=mapping_rules or {},
                sheet_layout=sheet_layout,
                DAF_mode=DAF_mode,
                custom_mode=custom_mode,
                static_content=static_content or {},
                pricing_net_weight=pricing_net_weight,
                footer_data=footer_data or {}
            )

        self.data_source_type = self.context.data_source_type
        self.data_source = self.context.data_source
        self.mapping_rules = self.context.mapping_rules
        self.sheet_layout = self.context.sheet_layout
        self.DAF_mode = self.context.DAF_mode
        self.custom_mode = self.context.custom_mode
        self.static_content = self.context.static_content
        self.pricing_net_weight = self.context.pricing_net_weight
        self.footer_data = self.context.footer_data
        
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
        
        # Cached parsed rules
        self._parsed_rules = None
    
    def resolve(self) -> ResolvedTableData:
        """
        Main resolution method - transforms raw data into table-ready rows.
        
        Returns:
            ResolvedTableData model instance containing prepared rows.
        """
        # Lazy import to avoid circular dependency
        from ..footer import TableFooterMapper

        # Parse mapping rules first
        parsed = self._parse_mapping_rules()
        
        # Extract data for this specific table (if multi-table)
        table_data_source = extract_table_data(self.data_source, self.data_source_type)
        
        # Prepare data rows using data_preparer logic
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
        footer_mapper = TableFooterMapper(context=self.context)
        resolved_footer = footer_mapper.resolve(data_rows, num_data_rows)
        
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
        """
        Factory method to create TableDataMapper from bundle configs.
        """
        args = context_config.get('args')
        DAF_mode = args.DAF if args and hasattr(args, 'DAF') else False
        custom_mode = args.custom if args and hasattr(args, 'custom') else False
        
        invoice_data = context_config.get('invoice_data') or {}
        metadata = invoice_data.get('metadata')
        if metadata is None:
            raise TableDataMapperError("CRITICAL: Invoice 'metadata' is missing or null in the provided JSON data.")
            
        context = MappingContext.from_bundles(
            data_config=data_config,
            context_config={
                **context_config,
                'DAF_mode': DAF_mode,
                'custom_mode': custom_mode,
                'pricing_net_weight': metadata.get('pricing_net_weight', False)
            },
            layout_config=layout_config
        )

        return TableDataMapper(context=context)
