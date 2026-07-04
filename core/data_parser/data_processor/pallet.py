import logging
import re
from typing import List, Dict, Any

def normalize_pallet_count(table_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Normalizes ALL pallet count values to binary 1/0.
    """
    prefix = "[normalize_pallet_count]"
    pallet_key = 'col_pallet_count'
    
    if not table_data:
        return table_data
    
    has_pallet = any(row.get(pallet_key) is not None for row in table_data)
    if not has_pallet:
        return table_data
    
    is_xy_format = False
    for row in table_data:
        v = row.get(pallet_key)
        if v is not None and str(v).strip():
            if re.match(r'^\d+-\d+$', str(v).strip()):
                is_xy_format = True
            break
    
    normalized_count = 0
    
    if is_xy_format:
        logging.info(f"{prefix} Detected 'x-y' pallet format. Normalizing to 1/0...")
        last_pallet_num = None
        
        for row in table_data:
            raw_val = row.get(pallet_key)
            if raw_val is None or not str(raw_val).strip():
                row[pallet_key] = 0
                continue
            
            raw_str = str(raw_val).strip()
            match = re.match(r'^(\d+)-(\d+)$', raw_str)
            
            if match:
                pallet_num = int(match.group(1))
                if pallet_num != last_pallet_num:
                    row[pallet_key] = 1
                    last_pallet_num = pallet_num
                    normalized_count += 1
                else:
                    row[pallet_key] = 0
            else:
                try:
                    row[pallet_key] = 1 if int(float(raw_str)) >= 1 else 0
                    if row[pallet_key] == 1:
                        normalized_count += 1
                except (ValueError, TypeError):
                    row[pallet_key] = 0
    else:
        logging.info(f"{prefix} Normalizing plain pallet values to 1/0...")
        for row in table_data:
            raw_val = row.get(pallet_key)
            if raw_val is None or not str(raw_val).strip():
                row[pallet_key] = 0
                continue
            
            try:
                int_val = int(float(str(raw_val).strip()))
                row[pallet_key] = 1 if int_val >= 1 else 0
                if row[pallet_key] == 1:
                    normalized_count += 1
            except (ValueError, TypeError):
                row[pallet_key] = 0
    
    logging.info(f"{prefix} Normalized {len(table_data)} rows → {normalized_count} pallet boundaries detected.")
    return table_data


def format_pallet_counts_to_xy(
    processed_tables: List[List[Dict[str, Any]]],
    total_pallets: int
) -> None:
    """
    Formats col_pallet_count from 1/0 binary flags into "x-y" strings (e.g. "3-19")
    in-place across all tables.
    """
    pallet_col_id = 'col_pallet_count'
    starting_pallet_order = 0
    
    for table_data in processed_tables:
        if not isinstance(table_data, list):
            continue
            
        pallet_order = starting_pallet_order
        carry_value = 0
        table_pallet_count = 0
        
        for row in table_data:
            if not isinstance(row, dict):
                continue
            val = row.get(pallet_col_id, 0)
            try:
                numeric_val = int(float(val)) if val is not None else 0
            except (ValueError, TypeError):
                numeric_val = 0
                
            if numeric_val == 1:
                pallet_order += 1
                table_pallet_count += 1
                formatted_val = f"{pallet_order}-{total_pallets}"
                row['col_pallet_no'] = formatted_val
                carry_value = formatted_val
            else:
                row['col_pallet_no'] = carry_value if carry_value != 0 else ""
                
        starting_pallet_order += table_pallet_count
