import logging
import re
from typing import Dict, List, Any, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet

from core.blueprint_generator.schema import BlueprintSchema
from core.utils.loop_profiler import loop_profiler, tick
from .models import ZoneBoundaries

logger = logging.getLogger(__name__)

def get_cell_value(cell) -> str:
    """Safe string value from cell."""
    if cell.value is None:
        return ""
    return str(cell.value).strip()


class _BoundaryScanner:
    """Short-lived, stateful scanner to find the header row in a single worksheet."""
    
    def __init__(self, worksheet: Worksheet, header_mappings: Optional[Dict[str, str]] = None):
        self.logger = logger
        self.worksheet = worksheet
        self.header_mappings = header_mappings or {}
        self.normalized_mappings: set = set()
        self._load_mappings()

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
    def find_header_row_structural(self, max_rows: int = 50) -> Optional[int]:
        """
        Legacy Structural Header Detection (Fallback).
        Finds row that is "widest" (most columns) and "text-heavy".
        """
        candidates = []
        
        for row in range(1, min(self.worksheet.max_row + 1, max_rows)):
            tick("scanner._find_header_row", sub="structural_rows_scanned")
            cells = []
            max_col_idx = 0
            # Cap structural scan at 50 columns
            for col in range(1, min(self.worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
                tick("scanner._find_header_row", sub="structural_cells_checked")
                cell = self.worksheet.cell(row=row, column=col)
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

    def _load_mappings(self) -> None:
        """Pre-build normalized user mappings for O(1) exact match lookup."""
        self.normalized_mappings = set()
        for m in self.header_mappings:
            self.normalized_mappings.add("".join(m.lower().split()))

    def _is_header_cell(self, value: str) -> bool:
        """Check if a cell value matches header keywords (User config or system rules)."""
        # 1. Check user mapping
        clean_val = "".join(value.lower().split())
        if self.header_mappings:
            tick("scanner._find_header_row", sub="user_mapping_checks")
            if value in self.header_mappings or clean_val in self.normalized_mappings:
                return True
        
        # 2. Check system rules
        tick("scanner._find_header_row", sub="rules_keyword_lookups")
        if BlueprintSchema.get_column_by_keyword(value):
            return True
            
        return False

    def _score_row(self, row: int) -> Tuple[int, int]:
        """Compute match count and total text count for a row."""
        matches = 0
        text_count = 0
        
        # Cap header scan at 25 (Col Y)
        for col in range(1, min(self.worksheet.max_column + 1, BlueprintSchema.MAX_SCAN_COLUMN)):
            tick("scanner._find_header_row", sub="cells_checked")
            cell = self.worksheet.cell(row=row, column=col)
            value = get_cell_value(cell)
            if value:
                text_count += 1
                if self._is_header_cell(value):
                    matches += 1
                    
        return matches, text_count

    def _is_better_candidate(self, matches: int, text_count: int, best_row: Optional[int], max_score: Tuple[int, int]) -> bool:
        """Evaluate if the current row is a better header candidate than the previous best."""
        # 1. Validity Threshold: Row must have AT LEAST 3 Valid Matches (keyword hits).
        if matches < 3:
            return False
            
        if best_row is None:
            return True
            
        best_matches, best_text_count = max_score
        
        # 2. Rank primarily by matches (more keyword hits = stronger header candidate).
        if matches > best_matches:
            return True
            
        # 3. Tie-breaker: text_count (more occupied cells = wider row).
        # 4. Tie-breaker: Topmost row (implicit: we iterate top-down, strict >).
        if matches == best_matches and text_count > best_text_count:
            return True
            
        return False

    def _fallback_scan(self, max_rows: int) -> Optional[int]:
        """Attempt to find the header row using legacy structural fallback."""
        self.logger.warning("Header detection: No header row found meeting threshold (min 3 matches). Trying Legacy Structural Fallback...")
        fallback_row = self.find_header_row_structural(max_rows=max_rows)
        if fallback_row:
            self.logger.info(f"Header detection (Fallback): Found structural header at row {fallback_row}.")
            return fallback_row
            
        self.logger.warning("Header detection: Fallback also failed.")
        return None

    @loop_profiler.watch("scanner._find_header_row")
    def scan(self) -> Optional[int]:
        """Find the header row by looking for known column keywords."""
        max_scan_rows = 50
        # Score tuple: (matches, text_count)
        max_score = (0, 0)
        best_row = None
                 
        for row in range(1, min(self.worksheet.max_row + 1, max_scan_rows)):
            tick("scanner._find_header_row", sub="rows_scanned")
            matches, text_count = self._score_row(row)
            
            if self._is_better_candidate(matches, text_count, best_row, max_score):
                 max_score = (matches, text_count)
                 best_row = row
             
        if best_row:
             self.logger.info(f"Header detection: Found header at row {best_row} with Score(matches={max_score[0]}, text={max_score[1]}).")
             return best_row
             
        return self._fallback_scan(max_scan_rows)


class BoundaryDetector:
    """Detects boundaries of all zones (Zone 1, 2, 3) in Excel worksheets."""
    
    def __init__(self):
        self.logger = logger

    def find_header_row(self, worksheet: Worksheet, header_mappings: Optional[Dict[str, str]] = None) -> Optional[int]:
        """Delegate scanning to a short-lived, stateful scanner to keep operations clean and thread-safe."""
        scanner = _BoundaryScanner(worksheet, header_mappings)
        return scanner.scan()

    def detect_boundaries(self, worksheet: Worksheet, header_mappings: Optional[Dict[str, str]] = None,
                          mapping_config: Optional[Dict[str, Any]] = None, sheet_name: str = "Unknown") -> Optional[ZoneBoundaries]:
        header_row = self.find_header_row(worksheet, header_mappings)
        if not header_row:
            return None
            
        # Determine data_start_row
        has_multi_row = self._check_multi_row_header(worksheet, header_row)
        data_start_row = header_row + 1
        if has_multi_row:
            for merged in worksheet.merged_cells.ranges:
                if merged.min_row == header_row:
                    data_start_row = max(data_start_row, merged.max_row + 1)
                    
        # Determine max scan column dynamically based on header
        header_max_col = 0
        if header_row > 1:
            for row in worksheet.iter_rows(min_row=1, max_row=header_row, min_col=1, max_col=40):
                for cell in row:
                    if cell.value is not None and cell.column > header_max_col:
                        header_max_col = cell.column
        max_col = min(worksheet.max_column, 40, max(20, header_max_col + 1))

        start_scan = data_start_row
        end_scan = min(worksheet.max_row, header_row + 500)
        
        # 1. Primary Method: Formula Adjacency check
        footer_row = self._find_footer_row_by_formula_adjacency(worksheet, start_scan, end_scan, max_col)
        if footer_row:
            self.logger.info(f"    Footer detected via formula adjacency at row {footer_row}")
        
        # 2. Fallback 1: Text keyword match (using find_total_label_cell)
        if not footer_row:
            from core.blueprint_generator.utils.content_extractor import find_total_label_cell
            total_cell_result = find_total_label_cell(worksheet, start_scan, end_scan, mapping_config, self.logger, sheet_name)
            if total_cell_result:
                found_cell, is_exact = total_cell_result
                footer_row = found_cell.row
                self.logger.info(f"    Footer detected via keyword match at row {footer_row}")
                
        # 3. Fallback 2: Strict bottom-up TOTAL keyword scan
        if not footer_row:
            scan_limit_bottom_up = max(start_scan, worksheet.max_row - 500)
            from core.blueprint_generator.utils.content_extractor import _get_cell_value_safe
            for row in range(worksheet.max_row, scan_limit_bottom_up - 1, -1):
                tick("scanner.detect_boundaries_strict", sub="fallback_rows_scanned")
                for col in range(1, max_col + 1):
                    tick("scanner.detect_boundaries_strict", sub="fallback_cells_checked")
                    cell = worksheet.cell(row=row, column=col)
                    value = _get_cell_value_safe(worksheet, cell)
                    if value:
                        val_upper = value.upper().strip()
                        if val_upper in ["TOTAL", "TOTAL:", "TOTAL OF:", "TOTAL OF", "TOTAL：", "TOTAL AMOUNT", "TOTAL AMOUNT:", "TOTAL AMOUNT："] or val_upper.startswith("TOTAL OF") or val_upper.startswith("TOTAL AMOUNT"):
                            footer_row = row
                            self.logger.warning(f"    Using strict TOTAL keyword fallback at row {footer_row}.")
                            break
                if footer_row:
                    break
                    
        footer_end_row = None
        if footer_row:
            footer_end_row = self._detect_footer_end_row(worksheet, footer_row, max_col)
            self.logger.info(f"    Contiguous footer block end detected at row {footer_end_row}")
                    
        return ZoneBoundaries(
            header_row=header_row,
            data_start_row=data_start_row,
            footer_row=footer_row,
            footer_end_row=footer_end_row,
            max_col=max_col
        )

    def _find_footer_row_by_formula_adjacency(self, ws: Worksheet, start_row: int, end_row: int, max_col: int) -> Optional[int]:
        """Find the footer row by scanning for =SUM or =SUBTOTAL formula adjacency."""
        from core.blueprint_generator.utils.content_extractor import _get_cell_value_safe
        
        last_candidate = None
        for row in range(start_row, end_row + 1):
            tick("scanner.detect_boundaries_formula", sub="rows_scanned")
            formula_cols = []
            for col in range(1, max_col + 1):
                tick("scanner.detect_boundaries_formula", sub="cells_checked")
                cell = ws.cell(row=row, column=col)
                value = _get_cell_value_safe(ws, cell)
                if value and value.startswith("="):
                    upper_val = value.upper()
                    if "=SUM(" in upper_val or "=SUBTOTAL(" in upper_val:
                        formula_cols.append(col)
            if len(formula_cols) >= 2:
                # Check for adjacency
                has_adjacent = False
                for i in range(len(formula_cols) - 1):
                    if formula_cols[i + 1] - formula_cols[i] == 1:
                        has_adjacent = True
                        break
                if has_adjacent:
                    last_candidate = row
        return last_candidate

    def _detect_footer_end_row(self, ws: Worksheet, footer_row: int, max_col: int) -> int:
        """Scan downwards from footer_row to find the end of the contiguous footer block (e.g. addon rows)."""
        from core.blueprint_generator.utils.content_extractor import _get_cell_value_safe
        
        current_row = footer_row
        max_search_row = min(ws.max_row, footer_row + 15)  # Scan at most 15 rows down
        
        addon_keywords = [
            "LEATHER", "COW", "BUFFALO", "PALLET", "PALLETS", "WEIGHT", "NW", "GW", "KGS", "PCS", "CBM", 
            "NET", "GROSS", "TOTAL", "SUBTOTAL", "SUM"
        ]
        
        signature_keywords = [
            "SIGNATURE", "APPROVED", "PREPARED", "CHECKED", "RECEIVER", "MANAGER", "DIRECTOR",
            "CHOP", "STAMP", "COMPANY", "BANK", "BENEFICIARY", "ACCOUNT", "SWIFT"
        ]
        
        for row in range(footer_row + 1, max_search_row + 1):
            tick("scanner.detect_footer_end_row", sub="rows_scanned")
            # Check if row is empty
            row_has_content = False
            has_addon_keyword = False
            has_signature_keyword = False
            has_numeric_value = False
            
            for col in range(1, max_col + 1):
                tick("scanner.detect_footer_end_row", sub="cells_checked")
                cell = ws.cell(row=row, column=col)
                value = _get_cell_value_safe(ws, cell)
                if value is not None and str(value).strip():
                    row_has_content = True
                    val_str = str(value).strip()
                    val_upper = val_str.upper()
                    
                    # Check for addon keywords
                    if any(kw in val_upper for kw in addon_keywords) or val_str.startswith("="):
                        has_addon_keyword = True
                        
                    # Check for signature keywords (which mark the start of static wrapper)
                    if any(kw in val_upper for kw in signature_keywords):
                        has_signature_keyword = True
                        
                    # Check if the value is numeric
                    if not val_str.startswith("="):
                        try:
                            clean_num = re.sub(r'[^\d.]', '', val_str)
                            if clean_num:
                                float(clean_num)
                                has_numeric_value = True
                        except ValueError:
                            pass
            
            if not row_has_content:
                # If we hit an empty row, the contiguous footer block ends at the previous row
                break
                
            if has_signature_keyword:
                # If we hit a signature block keyword, the dynamic footer block has ended
                break
                
            # A row is identified as a footer block row if it contains keywords AND contains at least one numeric value
            if has_addon_keyword and has_numeric_value:
                current_row = row
            else:
                # If it doesn't meet the footer heuristics, it belongs to the static template wrapper
                break
                
        return current_row

    def _check_multi_row_header(self, worksheet: Worksheet, header_row: int) -> bool:
        """Check if there's a multi-row header structure."""
        for merged in worksheet.merged_cells.ranges:
            tick("scanner._check_multi_row_header", sub="merge_ranges_checked")
            if merged.min_row == header_row and merged.max_row > header_row:
                return True
        return False
