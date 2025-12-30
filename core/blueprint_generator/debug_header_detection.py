
import sys
from pathlib import Path
import openpyxl
import logging

# Add project root to path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.append(str(project_root))

from core.blueprint_generator.excel_scanner import ExcelLayoutScanner
from core.blueprint_generator.blueprint_generator import BlueprintGenerator

# Setup logging
logging.basicConfig(level=logging.INFO)

def create_test_file(filename="test_invoice.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice"
    
    # Row 1-4: Metadata
    ws.append(["Customer:", "ABC Corp"])
    ws.append(["Date:", "2023-01-01"])
    ws.append([])
    
    # Row 5: Headers (Standard)
    headers = ["P.O Nº", "Item No", "Description", "Quantity", "Unit Price", "Amount"]
    ws.append(headers)
    
    # Row 6-10: Data
    for i in range(5):
        ws.append([f"PO-00{i}", f"ITEM-{i}", f"Desc {i}", 10, 5.0, 50.0])
        
    wb.save(filename)
    print(f"Created {filename}")
    return filename

def run_test():
    filename = create_test_file()
    
    print("\n--- Testing BlueprintGenerator (which uses ExcelLayoutScanner) ---")
    try:
        generator = BlueprintGenerator()
        # Verify mapping config loaded
        print(f"Mapping config path: {generator.mapping_config_path}")
        print(f"Mapping config exists: {generator.mapping_config_path.exists()}")
        
        # Run scan
        config = generator._load_mapping_config()
        print(f"Loaded {len(config.get('header_text_mappings', {}).get('mappings', {}))} mappings")
        
        analysis = generator.scanner.scan_template(filename, mapping_config=config)
        
        print("\n--- Analysis Result ---")
        for sheet in analysis.sheets:
            print(f"Sheet: {sheet.name}")
            print(f"  Header Row: {sheet.header_row}")
            if sheet.header_row == 5:
                print("  SUCCESS: Detected correct header row (5)")
            else:
                print(f"  FAILURE: Detected wrong header row ({sheet.header_row})")
                
            print(f"  Columns: {[c.header for c in sheet.columns]}")
            
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_test()
