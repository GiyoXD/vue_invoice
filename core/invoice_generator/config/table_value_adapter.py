"""
Multi-Table Data Adapter

This adapter is responsible for preparing table-specific data for rendering.
It transforms raw invoice data into table-ready row dictionaries based on:
- Data source type (aggregation, DAF_aggregation, custom, processed_tables)
- Mapping rules (which columns get which data)
- Column configurations (formats, IDs, etc.)

This eliminates data preparation logic from builders and centralizes it
in a single, testable, reusable adapter component.

Pattern:
    BundledConfigLoader → BuilderConfigResolver → TableDataAdapter → Builder
"""

import logging
from typing import Any, Dict, List, Tuple, Union, Optional
from decimal import Decimal
import ast
import re

from ..data.data_preparer import (
    prepare_data_rows,
    parse_mapping_rules,
    _to_numeric,
    _apply_fallback
)

logger = logging.getLogger(__name__)


class TableDataAdapter:
    """
    Adapter for preparing table-specific data for rendering.
    
    This class takes raw invoice data and configuration, then produces
    table-ready row dictionaries with proper formatting, formulas, and
    static values applied.
    
    Responsibilities:
    - Extract correct data subset for the table
    - Apply mapping rules to transform data → columns
    - Handle static values and formulas
    - Apply DAF/custom mode transformations
    - Calculate pallet counts and metadata
    
    Usage:
        adapter = TableDataAdapter(
            data_source_type='aggregation',
            data_source=invoice_data['standard_aggregation_results'],
            mapping_rules=config['mappings'],
            header_info=header_builder_result,
            DAF_mode=False
        )
        
        table_data = resolver.resolve()
        # Returns: {
        #     'data_rows': List[Dict[int, Any]],  # Ready-to-write rows
        #     'pallet_counts': List[int],          # Pallet count per row
        #     'num_data_rows': int                 # Total rows from source
        # }
    """
    
    def __init__(
        self,
        data_source_type: str,
        data_source: Union[Dict, List, None],
        mapping_rules: Dict[str, Any],
        header_info: Dict[str, Any],
        DAF_mode: bool = False,
        custom_mode: bool = False,
        table_key: Optional[str] = None,
        static_content: Optional[Dict[str, Any]] = None,
        footer_data: Optional[Dict[str, Any]] = None
    ):
        """
        Initialize the table data resolver.
        
        Args:
            data_source_type: Type of data source ('aggregation', 'DAF_aggregation', 
                            'custom_aggregation', 'processed_tables')
            data_source: Raw data from invoice_data
            mapping_rules: Mapping rules from config (how data maps to columns)
            header_info: Header information with column_map and column_id_map
            DAF_mode: Whether DAF mode is active
            custom_mode: Whether Custom mode is active
            table_key: Optional table key for multi-table data sources
            static_content: Static content from layout_bundle (e.g., col_static values)
            footer_data: Pre-calculated footer data from data parser (table_totals + grand_total)
        """
        self.data_source_type = data_source_type
        self.data_source = data_source
        self.mapping_rules = mapping_rules
        self.header_info = header_info
        self.DAF_mode = DAF_mode
        self.custom_mode = custom_mode
        self.table_key = table_key
        self.static_content = static_content or {}
        self.footer_data = footer_data or {}
        
        # Extract helper maps from header_info
        self.column_id_map = header_info.get('column_id_map', {})
        self.column_map = header_info.get('column_map', {})
        self.parent_column_ids = header_info.get('parent_column_ids', [])
        
        # Build reverse map (index → header)
        self.idx_to_header_map = {v: k for k, v in self.column_map.items()}
        
        # Cached parsed rules
        self._parsed_rules = None
    
    def resolve(self) -> Dict[str, Any]:
        """
        Main resolution method - transforms raw data into table-ready rows.
        
        Returns:
            Dictionary containing:
            - data_rows: List of row dictionaries (col_index → value)
            - num_data_rows: Number of data rows from source
            - static_info: Static configuration (col1_index, num_static_labels, etc.)
        """
        # Parse mapping rules first
        parsed = self._parse_mapping_rules()
        
        # Extract data for this specific table (if multi-table)
        table_data_source = self._extract_table_data()
        
        # Prepare data rows using the existing data_preparer logic
        data_rows, pallet_counts, num_data_rows = prepare_data_rows(
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
            parent_column_ids=self.parent_column_ids
        )
        
        logger.debug(f"[DEBUG-RESOLVE] Parsed Rules: {parsed['dynamic_mapping_rules'].keys()}")
        logger.debug(f"[DEBUG-RESOLVE] column_id_map: {self.column_id_map}")
        logger.debug(f"[DEBUG-RESOLVE] idx_to_header: {self.idx_to_header_map}")
        logger.debug(f"[DEBUG-RESOLVE] Returned row 0 (if any): {data_rows[0] if data_rows else 'EMPTY'}")
        
        # Merge static content with data rows (not prepend as separate rows)
        # Static content from layout_bundle.content.static should be merged into the first N data rows
        if self.static_content and 'col_static' in self.static_content:
            static_values = self.static_content['col_static']
            static_col_idx = self.column_id_map.get('col_static')
            
            if static_values and static_col_idx and len(data_rows) > 0:
                # [Smart Feature] Resolve {col_desc_fallback} placeholder for dynamic static content
                desc_fallback_str = ""
                for rule_key, rule in parsed['dynamic_mapping_rules'].items():
                    if 'desc' in rule_key.lower() and isinstance(rule, dict):
                        fc = rule.get('fallback')
                        if isinstance(fc, dict):
                            if self.DAF_mode and 'daf' in fc: desc_fallback_str = fc['daf']
                            elif self.custom_mode and 'custom' in fc: desc_fallback_str = fc['custom']
                            elif 'standard' in fc: desc_fallback_str = fc['standard']
                        elif fc is not None:
                            desc_fallback_str = str(fc)
                        break
                
                # Merge static values into the first N data rows
                num_static_values = len(static_values)
                
                # [Smart Feature] Extend data_rows if we have more static values than data rows
                # This ensures col_static dictates the minimum span of rows for that item block
                while len(data_rows) < num_static_values:
                    # Append an empty row (using same column keys as existing rows to prevent errors)
                    empty_row = {col_idx: "" for col_idx in self.column_id_map.values()}
                    data_rows.append(empty_row)

                for i, static_value in enumerate(static_values):
                    # Intercept {col_desc_fallback} in the text
                    if isinstance(static_value, str) and "{col_desc_fallback}" in static_value:
                        static_value = static_value.replace("{col_desc_fallback}", str(desc_fallback_str))
                        
                    # Add static value to the existing (or newly extended) data row
                    data_rows[i][static_col_idx] = static_value
                
                logger.info(f"Merged {num_static_values} static values into {len(data_rows)} data rows")
        
        # Extract summaries if available in data source
        leather_summary = None
        weight_summary = None
        pallet_summary_total = None
        
        if isinstance(self.data_source, dict):
            leather_summary = self.data_source.get('leather_summary')
            # Look in footer_data if not found directly in data_source
            if not leather_summary and self.footer_data and 'add_ons' in self.footer_data:
                leather_summary = self.footer_data['add_ons'].get('leather_summary_addon')
            
            weight_summary = self.data_source.get('weight_summary')
            if not weight_summary and self.footer_data and 'add_ons' in self.footer_data:
                weight_summary = self.footer_data['add_ons'].get('weight_summary_addon')
        elif isinstance(self.data_source, list):
            # If data_source is a list (like multi_table), summaries are strictly in footer_data
            if self.footer_data and 'add_ons' in self.footer_data:
                leather_summary = self.footer_data['add_ons'].get('leather_summary_addon')
                weight_summary = self.footer_data['add_ons'].get('weight_summary_addon')
        
        # Resolve pallet count from pre-calculated footer_data (top-level, from data parser)
        # table_totals[index] for per-table footers, grand_total for overall total
        if self.footer_data and 'table_totals' in self.footer_data:
            table_totals = self.footer_data['table_totals']
            if isinstance(table_totals, list) and len(table_totals) > 0:
                # table_key is the zero-based index used by MultiTableProcessor
                tbl_idx = 0
                if self.table_key is not None and str(self.table_key).isdigit():
                    tbl_idx = int(self.table_key)
                
                if tbl_idx >= len(table_totals):
                    tbl_idx = 0  # Safe fallback
                    logger.warning(f"table_key '{self.table_key}' out of bounds for table_totals (len={len(table_totals)}), using index 0")
                
                tbl_footer = table_totals[tbl_idx]
                if 'col_pallet_count' in tbl_footer:
                    pallet_summary_total = int(tbl_footer['col_pallet_count'])
                    logger.info(f"Using pre-calculated pallet count from footer_data.table_totals[{tbl_idx}]: {pallet_summary_total}")
            elif isinstance(table_totals, dict):
                # Fallback for old {"1": {...}} format
                first_val = next(iter(table_totals.values()), {})
                if 'col_pallet_count' in first_val:
                    pallet_summary_total = int(first_val['col_pallet_count'])
        
        # Last-resort fallback: check data_source directly (legacy path)
        if pallet_summary_total is None and isinstance(self.data_source, dict):
            pallet_summary_total = self.data_source.get('pallet_summary_total')
            if pallet_summary_total is not None:
                logger.warning(f"Using legacy pallet_summary_total from data_source: {pallet_summary_total}")

        return {
            'data_rows': data_rows,
            'pallet_counts': pallet_counts,
            'num_data_rows': num_data_rows,
            'static_info': {
                'col1_index': parsed['col1_index'],
                'num_static_labels': parsed['num_static_labels'],
                'initial_static_col1_values': parsed['initial_static_col1_values'],
                'static_column_header_name': parsed['static_column_header_name'],
                'apply_special_border_rule': parsed['apply_special_border_rule']
            },
            'formula_rules': parsed['formula_rules'],
            'static_content': self.static_content,  # Pass through static content from layout_bundle
            'leather_summary': leather_summary,
            'weight_summary': weight_summary,
            'pallet_summary_total': pallet_summary_total
        }
    
    def _parse_mapping_rules(self) -> Dict[str, Any]:
        """Parse mapping rules using existing data_preparer logic."""
        if self._parsed_rules is None:
            # Use mapping rules directly (data_preparer now supports bundled format)
            self._parsed_rules = parse_mapping_rules(
                mapping_rules=self.mapping_rules,
                column_id_map=self.column_id_map,
                idx_to_header_map=self.idx_to_header_map
            )
        return self._parsed_rules
    
    def _extract_table_data(self) -> Union[Dict, List, None]:
        """
        Extract data for the specific table being processed.
        
        For multi-table data sources, this extracts the subset for table_key.
        For single-table sources, returns the full data source.
        
        Note: BuilderConfigResolver.get_data_bundle(table_key) already extracts
        the specific table's data, so in most cases this just returns data_source as-is.
        """
        if self.data_source is None:
            return None
        
        # For processed_tables_multi, BuilderConfigResolver already extracted the table
        # So data_source is already the table data (dict with column arrays like {'po': [...], 'item': [...]})
        # Just return it as-is
        if self.data_source_type in ['processed_tables', 'processed_tables_multi']:
            return self.data_source
        
        # For other types like aggregation, return as-is
        # FIX: Check for stringified tuple keys (JSON artifact) and convert back to tuples
        if isinstance(self.data_source, dict):
            new_data = {}
            for k, v in self.data_source.items():
                if isinstance(k, str) and k.startswith('(') and k.endswith(')'):
                    try:
                        # It's a stringified tuple!
                        # Clean up Decimal wrappers for literal_eval: "Decimal('1.2')" -> "1.2"
                        clean_k = re.sub(r"Decimal\((['\"])(.*?)\1\)", r"\2", k)
                        new_key = ast.literal_eval(clean_k)
                        new_data[new_key] = v
                    except (ValueError, SyntaxError):
                        # Not a valid tuple string, keep as is
                        new_data[k] = v
                else:
                    new_data[k] = v
            return new_data

        return self.data_source
    
    def _get_desc_col_idx(self) -> int:
        """Get the description column index."""
        desc_col_id = None
        
        # Try common description column IDs
        for possible_id in ['col_desc', 'col_description', 'description']:
            if possible_id in self.column_id_map:
                desc_col_id = possible_id
                break
        
        return self.column_id_map.get(desc_col_id, -1) if desc_col_id else -1
    
    @staticmethod
    def create_from_bundles(
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> 'TableDataAdapter':
        """
        Factory method to create TableDataAdapter from bundle configs.
        
        This is the recommended way to instantiate the resolver when using
        the BuilderConfigResolver pattern.
        
        Args:
            data_config: Data bundle from BuilderConfigResolver.get_data_bundle()
            context_config: Context bundle from BuilderConfigResolver.get_context_bundle()
            layout_config: Optional layout bundle from BuilderConfigResolver.get_layout_bundle()
        
        Returns:
            TableDataAdapter instance
        """
        # Determine DAF mode and Custom mode
        args = context_config.get('args')
        DAF_mode = args.DAF if args and hasattr(args, 'DAF') else False
        custom_mode = args.custom if args and hasattr(args, 'custom') else False
        
        # Extract static_content from layout_config if provided
        static_content = {}
        if layout_config:
            static_content = layout_config.get('static_content', {})
        
        return TableDataAdapter(
            data_source_type=data_config.get('data_source_type', 'aggregation'),
            data_source=data_config.get('data_source'),
            mapping_rules=data_config.get('mapping_rules', {}),
            header_info=data_config.get('header_info', {}),
            DAF_mode=DAF_mode,
            custom_mode=custom_mode,
            table_key=data_config.get('table_key'),
            static_content=static_content,
            footer_data=data_config.get('footer_data', {})
        )


class TableDataAdapterError(Exception):
    """Exception raised when table data resolution fails."""
    pass
