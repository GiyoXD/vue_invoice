from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

def get_actual_column_width(worksheet: Worksheet, col_index: int, colspan: int = 1) -> float:
    """
    Calculate the total width of a column or merged column span.
    Handles openpyxl's grouped column dimensions (<col min="1" max="5" .../>)
    where dimensions are stored under a single dictionary key.
    """
    total_width = 0.0
    
    # Pre-cache dimension ranges because openpyxl stores grouped columns 
    # (e.g. <col min="1" max="5" width="20"/>) under a single dict key.
    dim_ranges = list(worksheet.column_dimensions.values())
    
    for c in range(col_index, col_index + colspan):
        matching_dim = None
        for dim in dim_ranges:
            if dim.min <= c <= dim.max:
                matching_dim = dim
                break
        
        # 1. Explicit width
        if matching_dim and matching_dim.width is not None:
            total_width += matching_dim.width
        # 2. Sheet Default
        elif worksheet.sheet_format and worksheet.sheet_format.defaultColWidth is not None:
            total_width += worksheet.sheet_format.defaultColWidth
        # 3. Fallback
        else:
            total_width += 15.0
            
    return total_width
