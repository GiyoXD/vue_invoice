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
        
        # Get bundles for TableFooterBuilder
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
                'args': Any,
                'invoice_data': dict,
                'pallets': int,
                'all_sheet_configs': dict,
            }
        """
        base_context = {
            'sheet_name': self.sheet_name,
            'args': self.args,
            'invoice_data': self.invoice_data,
            'pallets': self.pallets,
            'all_sheet_configs': self.config_loader.get_raw_config().get('layout_bundle', {}),
        }
        
        
        # Aggregate pre-calculated summaries from footer_data.grand_total
        footer_data = self.invoice_data.get('footer_data', {}) if self.invoice_data else {}
        grand_total = footer_data.get('grand_total', {})
        
        if grand_total:
            total_net = float(grand_total.get('col_net', 0))
            total_gross = float(grand_total.get('col_gross', 0))
            total_pallets = int(grand_total.get('col_pallet_count', 0))
            
            if not grand_total.get('col_pallet_count'):
                logger.warning("⚠ No col_pallet_count in footer_data.grand_total. Pallet count will be 0.")

            summaries = {
                'total_net_weight': total_net,
                'total_gross_weight': total_gross,
                'total_pallets': total_pallets
            }
            base_context.update(summaries)
            logger.debug(f"Added global summaries to context: {summaries}")
        
        # Merge in any overrides and additional context
        base_context.update(self.context_overrides)
        base_context.update(additional_context)
        
        return base_context
    
    # _adapt_invoice_data_for_sheet removed as text replacements are disabled
    
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
        if table_key and isinstance(data_source, list):
            try:
                # Convert the incoming string table_key (e.g. "0", "1") to an integer index
                idx = int(table_key)
                if 0 <= idx < len(data_source):
                    data_source = data_source[idx]
            except ValueError:
                logger.warning(f"Resolver: Invalid table_key '{table_key}' for list data source.")
        elif table_key and isinstance(data_source, dict):
            # Fallback for older dictionary-based multi-table data (if any remains)
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
            'footer_data': self.invoice_data.get('footer_data', {}) if self.invoice_data else {},
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
        is_last_table: bool = False
    ) -> Tuple[Dict, Dict, Dict]:
        """
        Get all bundles needed for TableFooterBuilder.
        
        Args:
            sum_ranges: Cell ranges to sum in footer formulas
            pallet_count: Pallet count for this footer
            is_last_table: Whether this is the last table in multi-table mode
        
        Returns:
            (style_config, context_config, data_config) tuple
        """
        style_config = self.get_style_bundle()
        
        # Add footer-specific context
        context_config = self.get_context_bundle(
            pallet_count=pallet_count if pallet_count is not None else self.pallets,
            is_last_table=is_last_table
        )
        
        data_config = self.get_data_bundle()
        
        # Add footer-specific data
        data_config.update({
            'sum_ranges': sum_ranges or [],
            'footer_config': self._sheet_config.get('layout_config', {}).get('footer', {}),
            'DAF_mode': self.args.DAF if self.args else False,
            'custom_mode': self.args.custom if self.args and hasattr(self.args, 'custom') else False,
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
        # but need to be passed to TableFooterBuilder via FooterData
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
        
        Refactored Logic (v2.2):
        1. STRUCTURED LOOKUP: Checks `single_table` for aggregation types or `multi_table` for granular ones.
        2. STRICT LOOKUP: Checks if 'data_source_type' exists as a direct key in invoice_data.
        3. LEGACY FALLBACK: Checks the hardcoded 'type_mapping' for backward compatibility.
        """
        if not self.invoice_data:
            return {}
            
        # --- PATH -1: Override generic aggregation for Summary Packing List ---
        # If the blueprint generator marked the sheet as generic "aggregation" but it is a
        # Summary Packing List, explicitly upgrade it so it routes to manifest_by_pallet_per_po.
        normalized_sheet = self.sheet_name.strip().lower()
        if data_source_type == 'aggregation' and normalized_sheet == 'summary packing list':
            logger.info("Auto-upgrading data_source_type from 'aggregation' to 'summary_packing_list' based on sheet name")
            data_source_type = 'summary_packing_list'
        
        # --- PATH 0: STRUCTURED LOOKUP (The "Newest Way" - v2.3) ---
        single_table_group = self.invoice_data.get('single_table', {})
        multi_table_group = self.invoice_data.get('multi_table', {})
        
        # Determine if we should look in single_table or multi_table
        if data_source_type in ['processed_tables_multi', 'processed_tables', 'detail_packing_list']:
            if multi_table_group:
                logger.debug(f"Smart Resolver: Found multi_table data for '{data_source_type}'")
                return multi_table_group
        else:
            # Single table types (aggregation, DAF_aggregation, custom_aggregation, summary_packing_list)
            # 1. Flag Checks: If DAF/Custom mode is on, look for suffixed keys first inside `single_table`.
            if self.args:
                if getattr(self.args, 'DAF', False):
                    daf_key = f"{data_source_type}_DAF"
                    if daf_key in single_table_group:
                        logger.debug(f"Smart Resolver: DAF Mode ON. Found variant '{daf_key}' in single_table")
                        return single_table_group[daf_key]
                
                if getattr(self.args, 'custom', False):
                    custom_key = f"{data_source_type}_custom"
                    if custom_key in single_table_group:
                        logger.debug(f"Smart Resolver: Custom Mode ON. Found variant '{custom_key}' in single_table")
                        return single_table_group[custom_key]
                        
            # 2. Base Lookup in single_table: If no flag match (or flags off), use exact key.
            if data_source_type in single_table_group:
                logger.debug(f"Smart Resolver: Found strict match for data_source '{data_source_type}' in single_table")
                return single_table_group[data_source_type]
            
            # 3. Handle Special Fallbacks strictly within single_table
            if data_source_type == 'summary_packing_list':
                if 'manifest_by_pallet_per_po' in single_table_group:
                    logger.info("Summary packing list requested: Using 'manifest_by_pallet_per_po' from single_table group")
                    return single_table_group['manifest_by_pallet_per_po']
                    
            if data_source_type == 'aggregation' and self.args and getattr(self.args, 'custom', False):
                if 'aggregation_custom' in single_table_group:
                    logger.info("Custom mode active: Using 'aggregation_custom' from single_table group")
                    return single_table_group['aggregation_custom']
                    
            if data_source_type in ['aggregation', 'DAF_aggregation'] and self.args and getattr(self.args, 'DAF', False):
                if 'aggregation_DAF' in single_table_group:
                    logger.info("DAF mode active: Using 'aggregation_DAF' from single_table group")
                    return single_table_group['aggregation_DAF']

        # --- If not found in structured paths, return empty ---
        logger.warning(f"Resolver: Data source '{data_source_type}' not found in structured single_table or multi_table groups.")
        return {}
    
    def get_all_sheet_configs(self) -> Dict[str, Any]:
        """
        Get configurations for all sheets (for cross-sheet references).
        
        Returns:
            Dictionary of all sheet configurations from layout_bundle
        """
        return self.config_loader.get_raw_config().get('layout_bundle', {})
