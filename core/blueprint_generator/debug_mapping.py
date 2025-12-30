
import sys
import logging
from pathlib import Path

# Setup path to allow imports
current_dir = Path(__file__).parent
sys.path.append(str(current_dir.parents[2])) # Add core/.. to path

from core.blueprint_generator.excel_scanner import ExcelLayoutScanner

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("DebugMapping")

def debug_mappings():
    print("--- Starting Debug Mapping ---")
    
    analyzer = ExcelLayoutScanner()
    
    print(f"Total Mappings: {len(analyzer.HEADER_MAPPINGS)}")
    
    # Check for specific Chinese keys
    check_keys = ["金额", "净重", "备注", "amount", "total"]
    
    found_count = 0
    for k in check_keys:
        val = analyzer.HEADER_MAPPINGS.get(k.lower())
        status = "✅ Found" if val else "❌ Missing"
        print(f"Key '{k}': {status} -> {val}")
        if val: found_count += 1
        
    print("----------------------------")
    if found_count >= 5:
        print("SUCCESS: Mappings loaded correctly.")
    else:
        print("FAILURE: Some mappings are missing.")

if __name__ == "__main__":
    debug_mappings()
