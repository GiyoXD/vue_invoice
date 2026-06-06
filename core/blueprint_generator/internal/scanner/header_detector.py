import logging
import re
from typing import Dict, List, Any, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet

from core.blueprint_generator.schema import BlueprintSchema
from core.utils.loop_profiler import loop_profiler, tick

logger = logging.getLogger(__name__)

def get_cell_value(cell) -> str:
    """Safe string value from cell."""
    if cell.value is None:
        return ""
    return str(cell.value).strip()

class HeaderDetector:
    """Detects header row in Excel worksheets."""
    
    def __init__(self):
        self.logger = logger

    def is_potential_header_row(self, cells) -> bool:
        """
        Check if a row is a potential header row (Legacy Logic).
        Rejects rows that are primarily numeric (>30%).
        """
        if not cells:
            return False
            
        numeric_count = 0
        total_count = len(cells)
        
        for cell in cells:
            val = cell.value
            if val is None:
                continue
                
            is_numeric = False
            if isinstance(val, (int, float)):
                is_numeric = True
            else:
                s_val = str(val).strip()
                # Check for number format like -123.45
                if s_val and re.match(r'^-?\d+(\.\d+)?$', s_val):
                    is_numeric = True
            
            if is_numeric:
                numeric_count += 1
        
        # If more than 30% of cells are numeric, it's likely a data row
        return (numeric_count / total_count) <= 0.3

    @loop_profiler.watch("scanner._find_header_row_structural")
    def find_header_row_structural(self, worksheet: Worksheet, max_rows: int = 50) -> Optional[int]:
        """
        Legacy Structural Header Detection (Fallback).
        Finds row that is "widest" (most columns) and "text-heavy".
        """
        candidates = []
        
        for row in range(1, min(worksheet.max_row + 1, max_rows)):
            tick("scanner._find_header_row", sub="structural_rows_scanned")
            cells = []
            max_col_idx = 0
            # Cap structural scan at 50 columns
            for col in range(1, min(worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
                tick("scanner._find_header_row", sub="structural_cells_checked")
                cell = worksheet.cell(row=row, column=col)
                if cell.value is not None and str(cell.value).strip():
                    cells.append(cell)
                    max_col_idx = col
            
            if cells and self.is_potential_header_row(cells):
                candidates.append({
                    'row_num': row,
                    'max_col': max_col_idx,
                    'cell_count': len(cells)
                })
        
        if not candidates:
            self.logger.warning("Structural Fallback: No structural candidates found.")
            return None
            
        # Filter for widest rows (tolerance 2)
        max_width = max(c['max_col'] for c in candidates)
        width_tolerance = 2
        wide_candidates = [c for c in candidates if c['max_col'] >= (max_width - width_tolerance)]
        
        self.logger.info(f"Structural Fallback: Found {len(candidates)} candidates. Max Width: {max_width}. Wide Candidates: {[c['row_num'] for c in wide_candidates]}")
        
        if not wide_candidates:
            wide_candidates = candidates
            
        # Sort by cell_count desc, then row_num asc
        wide_candidates.sort(key=lambda x: (-x['cell_count'], x['row_num']))
        
        best = wide_candidates[0]
        self.logger.info(f"Structural Fallback: Selected Row {best['row_num']} (Cells={best['cell_count']}, MaxCol={best['max_col']})")
        
        return best['row_num']

    @loop_profiler.watch("scanner._find_header_row")
    def find_header_row(self, worksheet: Worksheet, mapping_config: Optional[Dict[str, Any]] = None) -> Tuple[Optional[int], List[Tuple[int, str]]]:
        """Find the header row by looking for known column keywords."""
        max_scan_rows = 50
        max_matches = 0
        # Score tuple: (matches, text_count)
        max_score = (0, 0)
        best_row = None
        best_header_cells = []
        
        # Pre-build normalized user mappings for O(1) exact match lookup
        normalized_mappings = set()
        mappings = {}
        if mapping_config:
             mappings = mapping_config.get('header_text_mappings', {}).get('mappings', {})
             for m in mappings:
                 normalized_mappings.add("".join(m.lower().split()))
                 
        for row in range(1, min(worksheet.max_row + 1, max_scan_rows)):
            tick("scanner._find_header_row", sub="rows_scanned")
            matches = 0
            text_count = 0
            header_cells = []
            
            has_content = False
            # Cap header scan at 25 (Col Y)
            for col in range(1, min(worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
                tick("scanner._find_header_row", sub="cells_checked")
                cell = worksheet.cell(row=row, column=col)
                value = get_cell_value(cell)
                if value:
                    has_content = True
                    text_count += 1
                    
                    # strong, rule-based match check
                    is_match = False
                    
                    # 1. Check user mapping
                    clean_val = "".join(value.lower().split())
                    if mapping_config:
                          tick("scanner._find_header_row", sub="user_mapping_checks")
                          # Fast O(1) set lookup for normalized match, or O(1) exact match dict lookup
                          if value in mappings or clean_val in normalized_mappings:
                              is_match = True
                    
                    # 2. Check system rules
                    if not is_match:
                        tick("scanner._find_header_row", sub="rules_keyword_lookups")
                        if BlueprintSchema.get_column_by_keyword(value):
                            is_match = True
                        
                    if is_match:
                        matches += 1
                        
                    header_cells.append((col, value))
            
            # HEADER ROW SELECTION ALGORITHM:
            # 1. Row must have AT LEAST 3 Valid Matches (keyword hits).
            # 2. Rank primarily by matches (more keyword hits = stronger header candidate).
            # 3. Tie-breaker: text_count (more occupied cells = wider row).
            # 4. Tie-breaker: Topmost row (implicit: we iterate top-down, strict >).
            
            # Validity Threshold
            if matches >= 3:
                 is_better = False
                 if best_row is None:
                     is_better = True
                 else:
                     (best_matches, best_text_count) = max_score
                     
                     if matches > best_matches:
                         is_better = True
                     elif matches == best_matches:
                         if text_count > best_text_count:
                             is_better = True
                             
                 if is_better:
                     max_score = (matches, text_count)
                     best_row = row
                     best_header_cells = header_cells
             
        if best_row:
             self.logger.info(f"Header detection: Found header at row {best_row} with Score(matches={max_score[0]}, text={max_score[1]}).")
        else:
             self.logger.warning("Header detection: No header row found meeting threshold (min 3 matches). Trying Legacy Structural Fallback...")
             # Fallback to Legacy Structural Detection
             fallback_row = self.find_header_row_structural(worksheet, max_rows=max_scan_rows)
             if fallback_row:
                 self.logger.info(f"Header detection (Fallback): Found structural header at row {fallback_row}.")
                 best_row = fallback_row
                 # Re-extract cells with CAP
                 best_header_cells = []
                 for col in range(1, min(worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
                     cell = worksheet.cell(row=fallback_row, column=col)
                     val = get_cell_value(cell)
                     if val:
                         best_header_cells.append((col, val))
             else:
                 self.logger.warning("Header detection: Fallback also failed.")
              
        return best_row, best_header_cells
