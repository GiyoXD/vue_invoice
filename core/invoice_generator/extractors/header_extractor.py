import logging
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

class HeaderExtractor:
    """
    Extracts semantic information (Company Name, Address) from the captured header state.
    """

    @staticmethod
    def extract(header_state: List[List[Dict[str, Any]]]) -> Dict[str, Any]:
        """
        Analyzes the header state to find company details.
        
        Heuristic:
        1. Company Name: Often the first non-empty cell in the first few rows, 
           possibly with larger font or bold styling.
        2. Address: Often found in rows immediately following the name.
        """
        info = {
            "consignee_address": None
        }

        if not header_state:
            return info

        # Consignee Extraction
        # Logic: Find "Consignee" row -> Capture until "Ship" row -> Filter out "Consignee" label
        consignee_start_row = -1
        ship_row = -1

        # 1. Find Start and End Rows
        for i, row in enumerate(header_state):
            row_text = " ".join([str(cell.get('value', '')) for cell in row if cell.get('value')]).upper()
            
            if consignee_start_row == -1 and "CONSIGNEE" in row_text:
                consignee_start_row = i
            
            if consignee_start_row != -1 and i > consignee_start_row and "SHIP" in row_text:
                ship_row = i
                break
        
        # 2. Extract Content
        if consignee_start_row != -1:
            consignee_lines = []
            end_row = ship_row if ship_row != -1 else len(header_state)
            
            for i in range(consignee_start_row, end_row):
                row = header_state[i]
                line_parts = []
                for cell in row:
                    val = cell.get('value')
                    if val and isinstance(val, str) and val.strip():
                        val_clean = val.strip()
                        # Skip the label itself if it's just "Consignee" or "Consignee :"
                        if "CONSIGNEE" in val_clean.upper():
                            # If cell contains "Consignee : Address", try to split
                            if ":" in val_clean and len(val_clean) > 15: # heuristic length
                                parts = val_clean.split(":", 1)
                                if len(parts) > 1 and parts[1].strip():
                                    line_parts.append(parts[1].strip())
                            continue
                        
                        line_parts.append(val_clean)
                
                if line_parts:
                    consignee_lines.append(" ".join(line_parts))
            
            if consignee_lines:
                info["consignee_address"] = consignee_lines

        return info
