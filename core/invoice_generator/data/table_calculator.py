import logging
from typing import Any, Dict, List, Optional, Tuple, Union
from decimal import Decimal

from ..styling.models import FooterData
from ..utils.math_utils import safe_float_convert, safe_int_convert

logger = logging.getLogger(__name__)

class TableCalculator:
    """
    Calculates summary data (weights, pallets, leather types) from resolved table data.
    
    This class extracts business logic from the DataTableBuilder, allowing for
    separation of calculation and rendering.
    """
    
    def __init__(self, header_info: Dict[str, Any]):
        """
        Initialize the calculator.
        
        Args:
            header_info: Header information with column maps.
        """
        self.header_info = header_info
        self.col_id_map = header_info.get('column_id_map', {})
        self.idx_to_id_map = {v: k for k, v in self.col_id_map.items()}
        
        # Initialize summaries
        self.leather_summary = {
            'BUFFALO': {'pallet_count': 0},
            'COW': {'pallet_count': 0}
        }
        self.weight_summary = {
            'net': 0.0,
            'gross': 0.0
        }
        self.total_pallets = 0

    def calculate(self, resolved_data: Dict[str, Any]) -> FooterData:
        """
        Perform all calculations on the provided data.
        
        Args:
            resolved_data: The data prepared by TableDataAdapter.
            
        Returns:
            FooterData object containing all calculated summaries.
        """
        data_rows = resolved_data.get('data_rows', [])
        pallet_counts = resolved_data.get('pallet_counts', [])
        
        # Check for pre-calculated summaries from data_parser
        use_precalc_leather = False
        use_precalc_weight = False
        
        if 'leather_summary' in resolved_data and resolved_data['leather_summary']:
            self.leather_summary = resolved_data['leather_summary']
            use_precalc_leather = True
            
        if 'weight_summary' in resolved_data and resolved_data['weight_summary']:
            self.weight_summary = resolved_data['weight_summary']
            use_precalc_weight = True
            
        if 'pallet_summary_total' in resolved_data and resolved_data['pallet_summary_total'] is not None:
            self.total_pallets = int(resolved_data['pallet_summary_total'])
        else:
            # Fallback calculation
            self.total_pallets = sum(self._parse_pallet_count(p) for p in pallet_counts)
        
        # Process each row only if needed
        if not use_precalc_leather or not use_precalc_weight:
            for i, row_data in enumerate(data_rows):
                if not use_precalc_weight:
                    self._update_weight_summary(row_data)
                if not use_precalc_leather:
                    self._update_leather_summary(row_data, i, pallet_counts)
            
        # Determine row indices (logic moved from DataTableBuilder)
        num_columns = self.header_info.get('num_columns', 0)
        data_writing_start_row = self.header_info.get('second_row_index', 0) + 1
        actual_rows_to_process = len(data_rows)
        
        data_start_row = data_writing_start_row
        data_end_row = data_start_row + actual_rows_to_process - 1 if actual_rows_to_process > 0 else data_start_row - 1
        footer_row_final = data_end_row + 1
        
        return FooterData(
            footer_row_start_idx=footer_row_final,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            total_pallets=self.total_pallets,
            leather_summary=self.leather_summary,
            weight_summary=self.weight_summary
        )

    def _update_weight_summary(self, row_data: Dict[int, Any]):
        """Updates the running totals for Net and Gross weight."""
        net_col_idx = self.col_id_map.get('col_net_weight') or self.col_id_map.get('col_net')
        gross_col_idx = self.col_id_map.get('col_gross_weight') or self.col_id_map.get('col_gross')
        
        if net_col_idx and net_col_idx in row_data:
            self.weight_summary['net'] += safe_float_convert(row_data[net_col_idx])
                
        if gross_col_idx and gross_col_idx in row_data:
            self.weight_summary['gross'] += safe_float_convert(row_data[gross_col_idx])

    def _update_leather_summary(self, row_data: Dict[int, Any], row_index: int, pallet_counts: List[Any]):
        """Updates the running totals for Buffalo and Cow leather."""
        desc_col_idx = self.col_id_map.get('col_desc')
        if not desc_col_idx:
            return

        description = str(row_data.get(desc_col_idx, "")).upper()
        
        if "BUFFALO" in description:
            target_type = 'BUFFALO'
        else:
            target_type = 'COW'
            
        if target_type:
            # Add pallet count for this row
            if row_index < len(pallet_counts):
                self.leather_summary[target_type]['pallet_count'] += self._parse_pallet_count(pallet_counts[row_index])
            
            # Sum numeric columns
            for col_idx, value in row_data.items():
                col_id = self.idx_to_id_map.get(col_idx)
                if not col_id or col_id == 'col_desc':
                    continue
                
                num_val = safe_float_convert(value)

                if col_id not in self.leather_summary[target_type]:
                    self.leather_summary[target_type][col_id] = 0
                    
                self.leather_summary[target_type][col_id] += num_val

    def _parse_pallet_count(self, pallet_str: Any) -> int:
        """
        Parses a pallet string (like '1-25' or '1~5-25') to determine how many pallets it represents.
        Returns 1 as default if parsing fails but it's a non-empty string.
        """
        if pallet_str is None or pallet_str == "":
            return 0
            
        pallet_str = str(pallet_str).strip()
        
        # Try direct int conversion first
        val = safe_int_convert(pallet_str, default=-1)
        if val != -1:
            return val
            
        # Parse strings like "1-25" (Pallet 1 of 25 -> 1 pallet)
        # Or "1~5-25" (Pallets 1 through 5 of 25 -> 5 pallets)
        # Or "1,2-25" (Pallets 1 and 2 of 25 -> 2 pallets)
        # Or "1~5" (Pallets 1 to 5)
        
        # Strip the "-{total}" part if present (e.g., "-25")
        if "-" in pallet_str:
            parts = pallet_str.split("-")
            # If there are multiple dashes or it's just X-Y, we assume the part before the LAST dash is the identifier
            identifier = "-".join(parts[:-1]) 
        else:
            identifier = pallet_str
            
        identifier = identifier.strip()
        
        # Count pallets in identifier
        if "~" in identifier:
            range_parts = identifier.split("~")
            if len(range_parts) == 2:
                try:
                    start = int(range_parts[0].strip())
                    end = int(range_parts[1].strip())
                    return abs(end - start) + 1
                except ValueError:
                    pass
        elif "," in identifier or "&" in identifier:
            # Count the items
            separators = [",", "&"]
            for sep in separators:
                if sep in identifier:
                    return len([x for x in identifier.split(sep) if x.strip()])
                    
        # Default: if it's a string like "1" or "A", it represents 1 pallet
        return 1
