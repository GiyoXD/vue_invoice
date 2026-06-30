import logging
from typing import Any, Dict
from core.system_config import ConfigurationError

logger = logging.getLogger(__name__)

def resolve_header_info(layout_config: Dict[str, Any], args: Any = None) -> Dict[str, Any]:
    """
    Construct header_info from layout_bundle.structure.
    
    Transforms bundled config format into the header_info structure builders expect.
    Handles both simple columns and parent columns with children (for colspan headers).
    
    Args:
        layout_config: The layout configuration for the sheet
        args: CLI arguments or options
    
    Returns:
        {
            'second_row_index': int,
            'column_map': {header_name: col_index},
            'column_id_map': {col_id: col_index},
            'num_columns': int,
            'column_formats': {col_id: format_string},
            'column_colspan': {col_id: colspan}
        }
    """
    structure = layout_config.get('structure', {})
    columns = structure.get('columns', [])
    header_row = structure.get('header_row')
    
    if header_row is None:
         raise ConfigurationError("CRITICAL: No 'header_row' found in structure config. Builders cannot determine header placement.")
    
    # Filter columns based on DAF/custom mode flags
    DAF_mode = args.DAF if args and hasattr(args, 'DAF') else False
    custom_mode = args.custom if args and hasattr(args, 'custom') else False
    
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
    parent_column_ids = []  # Track parent columns that have children
    
    current_idx = 1
    
    for col_def in filtered_columns:
        col_id = col_def.get('id', f'col_{current_idx}')
        header = col_def.get('header', '')
        fmt = col_def.get('format')
        colspan = col_def.get('colspan', 1)
        children = col_def.get('children', [])
        
        # If column has children, process each child
        if children:
            parent_column_ids.append(col_id)
            # Parent column gets its own entry (for merged cell reference)
            column_map[header] = current_idx
            column_id_map[col_id] = current_idx
            
            # Parent column itself should not be horizontally merged in data rows
            column_colspan[col_id] = 1
            
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
        'column_colspan': column_colspan,  # Colspan info for automatic merging
        'parent_column_ids': parent_column_ids
    }
