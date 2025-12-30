# invoice_generator/config/builder_config_resolver.py
"""
Builder Config Resolver

This resolver sits between the BundledConfigLoader and the individual builders.
It extracts exactly what each builder needs from the configuration bundles,
providing clean, builder-specific argument dictionaries.

Pattern:
    BundledConfigLoader → BuilderConfigResolver → Builder
    
The resolver prevents builders from needing to understand the full config structure.
Each builder gets only its required arguments in the expected format.
"""

import logging
from typing import Any, Dict, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet



logger = logging.getLogger(__name__)


class BuilderConfigResolver:
    """
    Resolves and prepares configuration bundles for specific builders.
    
    This class bridges the gap between the BundledConfigLoader's structure
    and the bundle arguments that each builder expects.
    
    Usage:
        resolver = BuilderConfigResolver(
            config_loader=config_loader,
            sheet_name="Invoice",
            worksheet=worksheet,
            args=cli_args,
            invoice_data=invoice_data,
            pallets=31
        )
        
        # Get bundles for HeaderBuilder
        header_bundles = resolver.get_header_bundles()
        
        # Get bundles for DataTableBuilder
        datatable_bundles = resolver.get_datatable_bundles(table_key="table_1")
        
        # Get bundles for FooterBuilder
        footer_bundles = resolver.get_footer_bundles(sum_ranges=ranges, pallet_count=31)
    """
    
    def __init__(
        self,
        config_loader,  # BundledConfigLoader instance
        sheet_name: str,
        worksheet: Worksheet,
        args=None,  # CLI arguments
        invoice_data: Optional[Dict[str, Any]] = None,
        pallets: int = 0,
        **context_overrides
    ):
        """
        Initialize the resolver with the config loader and runtime context.
        
        Args:
            config_loader: BundledConfigLoader instance with loaded config
            sheet_name: Name of the sheet being processed
            worksheet: The worksheet object
            args: CLI arguments (for DAF mode, custom mode, etc.)
            invoice_data: Invoice data dictionary
            pallets: Pallet count for the current context
            **context_overrides: Additional context values to override
        """
        self.config_loader = config_loader
        self.sheet_name = sheet_name
        self.worksheet = worksheet
        self.args = args
        self.invoice_data = invoice_data
        self.pallets = pallets
        self.context_overrides = context_overrides
        
        # Cache the full sheet config
        self._sheet_config = config_loader.get_sheet_config(sheet_name)
    
    # ========== Bundle Preparation Methods ==========
    
    def get_style_bundle(self) -> Dict[str, Any]:
        """
        Get the style bundle for builders.
        
        Returns:
            {
                'styling_config': StylingConfigModel or dict
            }
        
        NOTE: If sheet config has 'columns' and 'row_contexts' (new format),
        those are passed directly as the styling_config.
        """
        # Check if using new format (columns + row_contexts at sheet level)
        if 'columns' in self._sheet_config and 'row_contexts' in self._sheet_config:
            # New format: return entire sheet config as styling_config
            return {
                'styling_config': {
                    'columns': self._sheet_config['columns'],
                    'row_contexts': self._sheet_config['row_contexts']
                }
            }
        else:
            # Old format: look for nested styling_config key
            return {
                'styling_config': self._sheet_config.get('styling_config', {})
            }
    
    def get_context_bundle(self, table_key: Optional[str] = None, **additional_context) -> Dict[str, Any]:
        """
        Get the context bundle for builders.
        
        Args:
            table_key: Optional table key for multi-table scenarios (e.g., '1', '2')
            **additional_context: Additional context to merge in
        
        Returns:
            {
                'sheet_name': str,
                'args': CLI args,
                'invoice_data': dict,  # Adapted invoice_data with normalized structure for text replacements
                'pallets': int,
                'all_sheet_configs': dict,  # For cross-sheet references
                ... (any additional context)
            }
        """
        # Adapt invoice_data to normalize data paths for text replacements
        adapted_invoice_data = self._adapt_invoice_data_for_sheet(table_key)
        
        base_context = {
            'sheet_name': self.sheet_name,
            'args': self.args,
            'invoice_data': adapted_invoice_data,  # Use adapted data with normalized paths
            'pallets': self.pallets,
            'all_sheet_configs': self.config_loader.get_raw_config().get('layout_bundle', {}),
        }
        
        # Add processed_tables_data separately for features that need raw data (e.g., weight_summary)
        if self.invoice_data and 'processed_tables_data' in self.invoice_data:
            base_context['processed_tables_data'] = self.invoice_data['processed_tables_data']
            
            # Aggregate pre-calculated summaries from data_parser
            # This replaces GlobalSummaryCalculator with a lighter-weight aggregation
            total_net = 0.0
            total_gross = 0.0
            total_pallets = 0
            
            for table_data in self.invoice_data['processed_tables_data'].values():
                # Check for footer_data first
                footer_data = table_data.get('footer_data', table_data)

                # Aggregate weights
                if 'weight_summary' in footer_data:
                    ws = footer_data['weight_summary']
                    total_net += float(ws.get('net', 0.0))
                    total_gross += float(ws.get('gross', 0.0))
                
                # Aggregate pallets
                if 'pallet_summary_total' in footer_data:
                    total_pallets += int(footer_data['pallet_summary_total'])
                elif 'pallet_count' in table_data:
                    # Fallback if summary missing
                    for p in table_data['pallet_count']:
                        if p is not None:
                            try:
                                total_pallets += int(float(p))
                            except (ValueError, TypeError):
                                pass

            summaries = {
                'total_net_weight': total_net,
                'total_gross_weight': total_gross,
                'total_pallets': total_pallets
            }
            
            # Add calculated summaries to context
            base_context.update(summaries)
            logger.debug(f"Added global summaries to context: {summaries}")
        
        # Merge in any overrides and additional context
        base_context.update(self.context_overrides)
        base_context.update(additional_context)
        
        return base_context
    
    def _adapt_invoice_data_for_sheet(self, table_key: Optional[str] = None) -> Dict[str, Any]:
        """
        Adapt invoice_data to provide normalized data paths for text replacements.
        
        This ensures all sheets can use the same replacement rule paths like:
        ["processed_tables_data", "1", "col_inv_no", 0]
        
        For sheets using 'aggregation' or 'DAF_aggregation', we create a synthetic
        processed_tables_data structure pointing to the first row of actual processed_tables.
        This allows Invoice and Contract sheets to access metadata (inv_no, inv_date, inv_ref)
        using the same data paths as Packing list.
        
        Args:
            table_key: Table key for multi-table sheets (e.g., '1', '2')
        
        Returns:
            Adapted invoice_data dict with normalized structure
        """
        if not self.invoice_data:
            return {}
        
        # Get the data source type for this sheet
        data_source_type = self._sheet_config.get('data_source', 'aggregation')
        
        # If already using processed_tables_data/multi, return as-is (already has correct structure)
        if data_source_type in ['processed_tables_multi', 'processed_tables']:
            return self.invoice_data
        
        # For aggregation-based sheets (Invoice, Contract), ensure they can access metadata
        # Copy the invoice_data (shallow copy to avoid modifying original)
        adapted_data = dict(self.invoice_data)
        
        # Transformation logic for aggregation-based sheets (Invoice, Contract)
        # They need metadata (inv_no, inv_date, inv_ref) to be available in a predictable location
        # for text replacement rules.
        
        # 1. Preferred: Check 'invoice_info' (New Standard)
        # If invoice_info exists, we don't need to synthesize processed_tables_data if text rules use invoice_info.
        # However, to be safe for legacy rules that might still look at processed_tables_data['1'], 
        # we can synthesize it OR just return as is if we trust rules are updated.
        # Given we updated text_replacement_rules to fallback, we can just return adapted_data 
        if 'invoice_info' in self.invoice_data:
             logger.debug(f"Sheet {self.sheet_name} will use top-level 'invoice_info' for metadata.")
             return adapted_data

        # 2. Legacy Fallback: processed_tables_data['1']
        # Check if processed_tables_data exists with metadata fields
        if 'processed_tables_data' not in self.invoice_data:
            logger.warning(f"No processed_tables_data found for sheet {self.sheet_name}, text replacements for JFINV/JFREF/JFTIME will not work")
            return adapted_data
        
        proc_tables = self.invoice_data['processed_tables_data']
        source_table_key = table_key or '1'  # Default to table '1' for metadata
        
        if source_table_key not in proc_tables:
            logger.warning(f"Table key '{source_table_key}' not found in processed_tables_data for sheet {self.sheet_name}")
            return adapted_data
        
        # The structure is already correct - processed_tables_data['1'] has col_inv_no, col_inv_date, col_inv_ref
        logger.debug(f"Sheet {self.sheet_name} (data_source={data_source_type}) will use processed_tables_data['{source_table_key}'] for metadata (col_inv_no, col_inv_date, col_inv_ref)")
        
        return adapted_data
    
    def get_layout_bundle(self) -> Dict[str, Any]:
        """
        Get the layout bundle for builders.
        
        Returns:
            {
                'sheet_config': dict,  # Layout configuration
                'blanks': dict,  # Blank row configs
                'static_content': dict,  # Static content configs
                'merge_rules': dict,  # Cell merge rules
                ...
            }
        """
        layout_config = self._sheet_config.get('layout_config', {})
        
        # Extract static_content from the 'content' section
        content_section = layout_config.get('content', {})
        static_section = content_section.get('static', {})
        
        return {
            'sheet_config': layout_config,
            'blanks': layout_config.get('blanks', {}),
            'static_content': static_section,  # Extract from content.static
            'merge_rules': layout_config.get('merge_rules', {}),
        }
    
    def get_data_bundle(self, table_key: Optional[str] = None) -> Dict[str, Any]:
        """
        Get the data bundle for builders.
        
        This bundles BOTH config (rules/structure) AND data (from invoice_data).
        
        Args:
            table_key: Optional table key for multi-table scenarios
        
        Returns:
            {
                'data_source': invoice data subset (from JSON file),
                'data_source_type': str,
                'header_info': dict (constructed from layout_bundle.structure),
                'mapping_rules': dict (from layout_bundle.data_flow.mappings),
                ...
            }
        """
        layout_config = self._sheet_config.get('layout_config', {})
        
        # Extract data source type from config
        data_source_type = self._sheet_config.get('data_source', 'aggregation')
        
        # Get actual data from invoice_data (JSON file)
        data_source = self._get_data_source_for_type(data_source_type)
        
        # For multi-table processing, extract the specific table's data
        # IMPORTANT: Only extract if data_source contains multiple tables (not pre-filtered)
        # Check if we have all_tables_data vs single_table_data:
        # - If data_source has keys like '1', '2', '3', extract the specific table
        # - If data_source is already a single table dict with keys like 'po', 'item', skip extraction
        if table_key and isinstance(data_source, dict):
            # Check if this looks like multi-table data (has numeric string keys)
            has_table_keys = any(str(k).isdigit() for k in data_source.keys())
            
            if has_table_keys:
                # Extract the specific table
                data_source = data_source.get(str(table_key), {})
        
        # Construct header_info from layout_bundle.structure
        header_info = self._construct_header_info(layout_config)
        
        # Extract mapping rules from layout_bundle.data_flow.mappings
        mapping_rules = layout_config.get('data_flow', {}).get('mappings', {})
        
        return {
            'data_source': data_source,
            'data_source_type': data_source_type,
            'header_info': header_info,
            'mapping_rules': mapping_rules,
            'table_key': table_key,
        }
    
    # ========== Builder-Specific Bundle Methods ==========
    
    def get_header_bundles(self) -> Tuple[Dict, Dict, Dict]:
        """
        Get all bundles needed for HeaderBuilder.
        
        Returns:
            (style_config, context_config, layout_config) tuple
        """
        style_config = self.get_style_bundle()
        context_config = self.get_context_bundle()
        layout_config = self.get_layout_bundle()
        
        return style_config, context_config, layout_config
    
    def get_datatable_bundles(self, table_key: Optional[str] = None) -> Tuple[Dict, Dict, Dict, Dict]:
        """
        Get all bundles needed for DataTableBuilder.
        
        Args:
            table_key: Optional table key for multi-table scenarios
        
        Returns:
            (style_config, context_config, layout_config, data_config) tuple
        """
        style_config = self.get_style_bundle()
        context_config = self.get_context_bundle()
        layout_config = self.get_layout_bundle()
        data_config = self.get_data_bundle(table_key=table_key)
        
        return style_config, context_config, layout_config, data_config
    
    def get_layout_bundles_with_data(self, table_key: Optional[str] = None) -> Tuple[Dict, Dict, Dict]:
        """
        Get bundles for LayoutBuilder (style, context, and merged layout+data).
        
        This is a convenience method that combines layout_config and data_config
        into a single bundle for LayoutBuilder, which expects data_source and
        mapping_rules to be in layout_config.
        
        Args:
            table_key: Optional table key for multi-table scenarios
        
        Returns:
            (style_config, context_config, merged_layout_config) tuple
            where merged_layout_config contains both layout structure AND data resolution
        """
        style_config = self.get_style_bundle()
        context_config = self.get_context_bundle(table_key=table_key)  # Pass table_key for data adaptation
        layout_config = self.get_layout_bundle()
        data_config = self.get_data_bundle(table_key=table_key)
        
        # Merge data_config into layout_config for LayoutBuilder
        # LayoutBuilder expects data_source, data_source_type, header_info, mapping_rules in layout_config
        merged_layout_config = {
            **layout_config,
            'data_source': data_config.get('data_source'),
            'data_source_type': data_config.get('data_source_type'),
            'header_info': data_config.get('header_info'),
            'mapping_rules': data_config.get('mapping_rules'),
        }
        
        return style_config, context_config, merged_layout_config
    
    def get_table_data_resolver(self, table_key: Optional[str] = None):
        """
        Create a TableDataAdapter for preparing table-specific data.
        
        This method provides a high-level interface to data preparation logic,
        eliminating the need for builders to handle data transformation directly.
        
        Args:
            table_key: Optional table key for multi-table scenarios
        
        Returns:
            TableDataAdapter instance
        
        Example:
            resolver = BuilderConfigResolver(...)
            table_data_resolver = resolver.get_table_data_resolver(table_key='1')
            table_data = table_data_resolver.resolve()
            
            # table_data contains:
            # - data_rows: Ready-to-write row dictionaries
            # - pallet_counts: Pallet counts per row
            # - dynamic_desc_used: Metadata
            # - static_info: Column 1 static values, etc.
            # - static_content: Static content from layout_bundle (e.g., col_static)
        """
        from .table_value_adapter import TableDataAdapter
        
        data_config = self.get_data_bundle(table_key=table_key)
        context_config = self.get_context_bundle()
        layout_config = self.get_layout_bundle()
        
        return TableDataAdapter.create_from_bundles(
            data_config=data_config,
            context_config=context_config,
            layout_config=layout_config
        )
    
    def get_footer_bundles(
        self,
        sum_ranges: Optional[list] = None,
        pallet_count: Optional[int] = None,
        is_last_table: bool = False,
        dynamic_desc_used: bool = False
    ) -> Tuple[Dict, Dict, Dict]:
        """
        Get all bundles needed for FooterBuilder.
        
        Args:
            sum_ranges: Cell ranges to sum in footer formulas
            pallet_count: Pallet count for this footer
            is_last_table: Whether this is the last table in multi-table mode
            dynamic_desc_used: Whether dynamic description was used
        
        Returns:
            (style_config, context_config, data_config) tuple
        """
        style_config = self.get_style_bundle()
        
        # Add footer-specific context
        context_config = self.get_context_bundle(
            pallet_count=pallet_count if pallet_count is not None else self.pallets,
            is_last_table=is_last_table,
            dynamic_desc_used=dynamic_desc_used
        )
        
        data_config = self.get_data_bundle()
        
        # Add footer-specific data
        data_config.update({
            'sum_ranges': sum_ranges or [],
            'footer_config': self._sheet_config.get('layout_config', {}).get('footer', {}),
            'DAF_mode': self.args.DAF if self.args else False,
        })
        
        return style_config, context_config, data_config
    
    def get_footer_data(
        self,
        footer_row_start_idx: int,
        data_start_row: int,
        data_end_row: int,
        pallet_count: Optional[int] = None,
        leather_summary: Optional[Dict] = None,
        weight_summary: Optional[Dict] = None
    ):
        """
        Create a fully populated FooterData object.
        
        This method normalizes the data flow by ensuring that weight_summary
        is always populated, either from the passed argument (local table data)
        or by calculating global defaults if missing/zero.
        
        Args:
            footer_row_start_idx: Row index where footer starts
            data_start_row: Start row of data
            data_end_row: End row of data
            pallet_count: Pallet count (defaults to self.pallets)
            leather_summary: Leather summary dict
            weight_summary: Weight summary dict (net/gross)
            
        Returns:
            FooterData object
        """
        from ..styling.models import FooterData
        
        # Use provided pallet count or default to context
        final_pallets = pallet_count if pallet_count is not None else self.pallets
        
        # Normalize weight summary
        final_weight_summary = {'net': 0.0, 'gross': 0.0}
        
        # 1. Try to use provided weight summary
        if weight_summary:
            final_weight_summary.update(weight_summary)
            
        # 2. If weights are zero, try to use global calculated weights
        # This handles cases like "Invoice" sheet where weights come from global context
        # but need to be passed to FooterBuilder via FooterData
        if final_weight_summary['net'] == 0 and final_weight_summary['gross'] == 0:
            # Ensure global summaries are calculated
            context = self.get_context_bundle()
            if 'total_net_weight' in context:
                final_weight_summary['net'] = context['total_net_weight']
            if 'total_gross_weight' in context:
                final_weight_summary['gross'] = context['total_gross_weight']
                
        return FooterData(
            footer_row_start_idx=footer_row_start_idx,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            total_pallets=final_pallets,
            leather_summary=leather_summary,
            weight_summary=final_weight_summary
        )
    
    # ========== Helper Methods ==========
    
    def _construct_header_info(self, layout_config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Construct header_info from layout_bundle.structure.
        
        Transforms bundled config format into the header_info structure builders expect.
        Handles both simple columns and parent columns with children (for colspan headers).
        
        Args:
            layout_config: The layout configuration for the sheet
        
        Returns:
            {
                'second_row_index': int,
                'column_map': {header_name: col_index},
                'column_id_map': {col_id: col_index},
                'num_columns': int,
                'column_formats': {col_id: format_string}
            }
        """
        structure = layout_config.get('structure', {})
        columns = structure.get('columns', [])
        header_row = structure.get('header_row', 1)
        
        # Filter columns based on DAF/custom mode flags
        DAF_mode = self.args.DAF if self.args and hasattr(self.args, 'DAF') else False
        custom_mode = self.args.custom if self.args and hasattr(self.args, 'custom') else False
        
        filtered_columns = []
        for col_def in columns:
            col_id = col_def.get('id', 'unknown')
            skip_in_daf = col_def.get('skip_in_daf', False)
            skip_in_custom = col_def.get('skip_in_custom', False)
            
            # Skip column if it has skip_in_daf flag and we're in DAF mode
            if DAF_mode and skip_in_daf:
                logger.info(f"Filtering out column '{col_id}' (skip_in_daf=True, DAF_mode=True)")
                continue
            # Skip column if it has skip_in_custom flag and we're in custom mode
            if custom_mode and skip_in_custom:
                logger.info(f"Filtering out column '{col_id}' (skip_in_custom=True, custom_mode=True)")
                continue
            filtered_columns.append(col_def)
        
        logger.debug(f"Column filtering: {len(columns)} total → {len(filtered_columns)} after filtering (DAF={DAF_mode}, custom={custom_mode})")
        
        # Build column_map (header_name -> index) and column_id_map (col_id -> index)
        column_map = {}
        column_id_map = {}
        column_formats = {}
        column_colspan = {}  # Track colspan for each column ID
        
        current_idx = 1
        
        for col_def in filtered_columns:
            col_id = col_def.get('id', f'col_{current_idx}')
            header = col_def.get('header', '')
            fmt = col_def.get('format')
            colspan = col_def.get('colspan', 1)
            children = col_def.get('children', [])
            
            # If column has children, process each child
            if children:
                # Parent column gets its own entry (for merged cell reference)
                column_map[header] = current_idx
                column_id_map[col_id] = current_idx
                
                # Parent column spans across all children
                column_colspan[col_id] = len(children)
                
                # Process each child column
                for child_def in children:
                    child_id = child_def.get('id', f'col_{current_idx}')
                    child_header = child_def.get('header', '')
                    child_fmt = child_def.get('format')
                    
                    column_map[child_header] = current_idx
                    column_id_map[child_id] = current_idx
                    
                    if child_fmt:
                        column_formats[child_id] = child_fmt
                    
                    # Children columns don't span (colspan=1)
                    column_colspan[child_id] = 1
                    
                    current_idx += 1
            else:
                # Simple column without children
                column_map[header] = current_idx
                column_id_map[col_id] = current_idx
                
                if fmt:
                    column_formats[col_id] = fmt
                
                # Store colspan for this column
                column_colspan[col_id] = colspan
                
                # Increment by colspan to skip the physical columns occupied by the merge
                # Example: col_static at column 1 with colspan=2 occupies columns 1-2,
                # so next column (col_po) should start at column 3
                current_idx += colspan
        
        # second_row_index represents the second row of the header (where data writing starts after)
        # If header is at row N, second row is at N+1
        return {
            'second_row_index': header_row + 1,
            'column_map': column_map,
            'column_id_map': column_id_map,
            'num_columns': current_idx - 1,  # Total columns processed
            'column_formats': column_formats,
            'column_colspan': column_colspan  # Colspan info for automatic merging
        }
    
    def _get_data_source_for_type(self, data_source_type: str) -> Any:
        """
        Extract the appropriate data source from invoice_data based on type.
        
        Args:
            data_source_type: Type of data source (aggregation, DAF_aggregation, etc.)
        
        Returns:
            The appropriate data source from invoice_data
        """
        if not self.invoice_data:
            return {}
        
        # Map data source types to invoice_data keys
        type_mapping = {
            'aggregation': 'standard_aggregation_results',
            'DAF_aggregation': 'standard_aggregation_results',  # DAF uses same data structure
            'custom_aggregation': 'custom_aggregation_results',
            'processed_tables_multi': 'processed_tables_data',
            'processed_tables': 'processed_tables_data',
        }
        
        data_key = type_mapping.get(data_source_type, 'standard_aggregation_results')
        return self.invoice_data.get(data_key, {})
    
    def get_all_sheet_configs(self) -> Dict[str, Any]:
        """
        Get configurations for all sheets (for cross-sheet references).
        
        Returns:
            Dictionary of all sheet configurations from layout_bundle
        """
        return self.config_loader.get_raw_config().get('layout_bundle', {})
