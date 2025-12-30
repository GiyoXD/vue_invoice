# Metadata Generation Module Extraction

This document details the code responsible for generating the metadata JSON file. This logic is self-contained and can be easily ported to other projects.

## Core Function: `generate_metadata`

This function takes various processing artifacts (status, execution time, invoice data) and serializes them into a structured JSON file alongside the output.

### Required Imports

```python
import json
import datetime
import logging
from pathlib import Path
from typing import Optional, Dict, List, Any
import argparse

logger = logging.getLogger(__name__)
```

### The Function Code

```python
def generate_metadata(
    output_path: Path, 
    status: str, 
    execution_time: float, 
    sheets_processed: List[str], 
    sheets_failed: List[str], 
    error_message: Optional[str] = None, 
    invoice_data: Optional[Dict] = None, 
    cli_args: Optional[argparse.Namespace] = None, 
    replacements_log: Optional[List[Dict]] = None, 
    header_info: Optional[Dict] = None
):
    """Generates a metadata JSON file for backend integration."""
    
    # 1. Basic Info
    metadata = {
        "status": status,
        "output_file": str(output_path),
        "timestamp": datetime.datetime.now().isoformat(),
        "execution_time": execution_time,
        "sheets_processed": sheets_processed,
        "sheets_failed": sheets_failed,
        "error_message": error_message
    }

    # 2. Config Info
    if cli_args:
        # Note: adjust attributes access based on your specific CLI args object
        metadata["config_info"] = {
            "daf_mode": getattr(cli_args, 'DAF', False),
            "custom_mode": getattr(cli_args, 'custom', False),
            "input_file": Path(arg.input_data_file).name if hasattr(cli_args, 'input_data_file') else "unknown",
            "config_dir": getattr(cli_args, 'configdir', "./configs")
        }

    # 3. Database Export (Packing List Items)
    if invoice_data and "processed_tables_data" in invoice_data:
        processed_data = invoice_data["processed_tables_data"]
        packing_list_items = []
        
        # Iterate through tables (e.g., "1", "2")
        for table_id, table_data in processed_data.items():
            # All lists in table_data should be same length
            row_count = len(table_data.get("po", []))
            
            for i in range(row_count):
                try:
                    item = {
                        "po": table_data.get("po", [])[i],
                        "item": table_data.get("item", [])[i],
                        "description": table_data.get("description", [])[i],
                        "pcs": table_data.get("pcs", [])[i],
                        "sqft": table_data.get("sqft", [])[i],
                        "pallet_count": table_data.get("pallet_count", [])[i],
                        "net": table_data.get("net", [])[i],
                        "gross": table_data.get("gross", [])[i],
                        "cbm": table_data.get("cbm", [])[i]
                    }
                    packing_list_items.append(item)
                except IndexError:
                    continue # Skip malformed rows

        # 4. Summary Statistics
        # robust calculation handling strings/integers
        total_pcs = sum(int(i["pcs"]) for i in packing_list_items if str(i["pcs"]).isdigit())
        total_sqft = sum(float(i["sqft"]) for i in packing_list_items if str(i["sqft"]).replace('.', '', 1).isdigit())
        total_pallets = sum(int(i["pallet_count"]) for i in packing_list_items if str(i["pallet_count"]).isdigit())
        
        metadata["database_export"] = {
            "summary": {
                "total_pcs": total_pcs,
                "total_sqft": total_sqft,
                "total_pallets": total_pallets,
                "item_count": len(packing_list_items)
            },
            "packing_list_items": packing_list_items
        }

    # 5. Invoice Terms (from text replacements)
    if replacements_log:
        terms_found = set()
        for entry in replacements_log:
            if entry.get("term"):
                terms_found.add(entry["term"])
        
        metadata["invoice_terms"] = {
            "detected_terms": list(terms_found),
            "replacements_detail": replacements_log
        }

    # 6. Header Info (Company Name, Address)
    if header_info:
        metadata["header_info"] = header_info

    # 7. Invoice Info (No, Date, Ref)
    if invoice_data and "processed_tables_data" in invoice_data:
        try:
            # Assuming data is in the first table "1"
            # Adjust logic if you have multiple invoices per file
            table_data = invoice_data["processed_tables_data"].get("1", {})
            
            # Helper to safely get the first value from a list
            def get_first(key):
                val = table_data.get(key)
                if isinstance(val, list) and val:
                    return val[0]
                return None

            metadata["invoice_info"] = {
                "inv_no": get_first("inv_no"),
                "inv_date": get_first("inv_date"),
                "inv_ref": get_first("inv_ref")
            }
        except Exception as e:
            # Don't fail metadata generation if extraction fails
            metadata["invoice_info_error"] = str(e)

    # Output file logic
    meta_path = output_path.parent / (output_path.name + ".meta.json")
    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2)
        logger.info(f"Metadata generated: {meta_path}")
    except Exception as e:
        logger.error(f"Failed to generate metadata: {e}")
```

## Data Structures Explanation

### 1. `invoice_data["processed_tables_data"]`
The code expects `invoice_data` to be a dictionary where `processed_tables_data` contains the raw columns extracted from the input.
Structure:
```json
{
  "processed_tables_data": {
    "1": {
      "po": ["PO123", "PO123"],
      "item": ["ITEM1", "ITEM2"],
      "pcs": [10, 20],
      ...
    }
  }
}
```

### 2. `replacements_log`
Used for tracking text replacements (like payment terms).
Expected structure: list of dicts.
```json
[
  { "term": "PAYMENT_TERMS", "original": "{PAYMENT}", "new": "Net 30" }
]
```

### 3. `header_info`
Contains company header details.
Expected structure: Key-value dict.
```json
{
  "company_name": "My Company",
  "address": "123 Street"
}
```

## How to Integrate

1.  Copy the imports and the `generate_metadata` function to your utility or main script.
2.  Ensure your data pipeline populates the `invoice_data` dictionary with the expected `processed_tables_data` structure if you want the "Database Export" feature.
3.  Call `generate_metadata` at the end of your processing workflow.

```python
# Example Usage
generate_metadata(
    output_path=Path("output/invoice.xlsx"),
    status="success",
    execution_time=1.5,
    sheets_processed=["Sheet1"],
    sheets_failed=[],
    invoice_data=my_data_dict
)
```
