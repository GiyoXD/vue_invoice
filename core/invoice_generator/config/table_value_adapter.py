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
        #     'dynamic_desc_used': bool,           # Whether dynamic descriptions were used
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
        table_key: Optional[str] = None,
        static_content: Optional[Dict[str, Any]] = None
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
            table_key: Optional table key for multi-table data sources
            static_content: Static content from layout_bundle (e.g., col_static values)
        """
        self.data_source_type = data_source_type
        self.data_source = data_source
        self.mapping_rules = mapping_rules
        self.header_info = header_info
        self.DAF_mode = DAF_mode
        self.table_key = table_key
        self.static_content = static_content or {}
        
        # Extract helper maps from header_info
        self.column_id_map = header_info.get('column_id_map', {})
        self.column_map = header_info.get('column_map', {})
        
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
            - pallet_counts: List of pallet counts per row (if applicable)
            - dynamic_desc_used: Whether dynamic descriptions were used
            - num_data_rows: Number of data rows from source
            - static_info: Static configuration (col1_index, num_static_labels, etc.)
        """
        # Parse mapping rules first
        parsed = self._parse_mapping_rules()
        
        # Extract data for this specific table (if multi-table)
        table_data_source = self._extract_table_data()
        
        # Prepare data rows using the existing data_preparer logic
        data_rows, pallet_counts, dynamic_desc_used, num_data_rows = prepare_data_rows(
            data_source_type=self.data_source_type,
            data_source=table_data_source,
            dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
            column_id_map=self.column_id_map,
            idx_to_header_map=self.idx_to_header_map,
            desc_col_idx=self._get_desc_col_idx(),
            num_static_labels=parsed['num_static_labels'],
            static_value_map=parsed['static_value_map'],
            DAF_mode=self.DAF_mode
        )
        
        # Merge static content with data rows (not prepend as separate rows)
        # Static content from layout_bundle.content.static should be merged into the first N data rows
        if self.static_content and 'col_static' in self.static_content:
            static_values = self.static_content['col_static']
            static_col_idx = self.column_id_map.get('col_static')
            
            if static_values and static_col_idx and len(data_rows) > 0:
                # Merge static values into the first N data rows
                num_static_values = len(static_values)
                
                for i, static_value in enumerate(static_values):
                    if i < len(data_rows):
                        # Add static value to the existing data row
                        data_rows[i][static_col_idx] = static_value
                
                logger.info(f"Merged {num_static_values} static values into first {num_static_values} data rows")
        
        # Extract summaries if available in data source
        leather_summary = None
        weight_summary = None
        pallet_summary_total = None
        
        if isinstance(self.data_source, dict):
            leather_summary = self.data_source.get('leather_summary')
            weight_summary = self.data_source.get('weight_summary')
            pallet_summary_total = self.data_source.get('pallet_summary_total')

        return {
            'data_rows': data_rows,
            'pallet_counts': pallet_counts,
            'dynamic_desc_used': dynamic_desc_used,
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
        # Determine DAF mode
        args = context_config.get('args')
        DAF_mode = args.DAF if args and hasattr(args, 'DAF') else False
        
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
            table_key=data_config.get('table_key'),
            static_content=static_content
        )


class TableDataAdapterError(Exception):
    """Exception raised when table data resolution fails."""
    pass
