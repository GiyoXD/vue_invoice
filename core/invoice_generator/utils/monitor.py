
import logging
import time
import datetime
import json
import traceback
from pathlib import Path
from typing import Optional, List, Any, Dict
import argparse

logger = logging.getLogger(__name__)

class GenerationMonitor:
    """
    Context manager to monitor invoice generation, track state, and GUARANTEE 
    metadata file generation upon exit (success or failure).
    """
    def __init__(self, output_path: Path, args: argparse.Namespace = None, input_data: Dict = None):
        self.output_path = Path(output_path)
        self.args = args
        self.input_data = input_data or {}
        
        self.start_time = None
        self.sheets_processed = []
        self.sheets_failed = []
        self.replacements_log = []
        self.header_info = {}
        
        self.status = "pending"
        self.error_message = None
        self.error_traceback = None

    def __enter__(self):
        self.start_time = time.time()
        logger.info(f"=== Generation Process Started ===")
        return self

    def log_success(self, sheet_name: str):
        self.sheets_processed.append(sheet_name)
        logger.info(f"Successfully processed sheet: {sheet_name}")

    def log_failure(self, sheet_name: str, error: Exception = None):
        self.sheets_failed.append(sheet_name)
        msg = f"Failed to process sheet {sheet_name}: {error}"
        logger.error(msg)
        if error:
            logger.debug(traceback.format_exc())

    def update_logs(self, replacements: List = None, header_info: Dict = None):
        if replacements:
            self.replacements_log.extend(replacements)
        if header_info:
            self.header_info.update(header_info)

    def __exit__(self, exc_type, exc_val, exc_tb):
        duration = time.time() - self.start_time
        
        # Determine status
        if exc_type:
            self.status = "fatal"
            self.error_message = str(exc_val)
            self.error_traceback = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
            logger.critical(f"Process crashed: {self.error_message}")
        elif self.sheets_failed:
            self.status = "partial_success" if self.sheets_processed else "error"
            self.error_message = f"Failed sheets: {self.sheets_failed}"
        else:
            self.status = "success"

        # Generate Metadata
        self._write_metadata(duration)
        
        # We generally want to propagate exceptions so the CLI/Orchestrator knows it failed
        return False 

    def _write_metadata(self, duration: float):
        """Write the metadata JSON file."""
        # Ensure output directory exists
        if not self.output_path.parent.exists():
            try:
                self.output_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                logger.error(f"Failed to create output directory for metadata: {e}")

        meta_path = self.output_path.parent / f"{self.output_path.stem}_metadata.json"
        
        # 1. Basic Info
        metadata = {
            "status": self.status,
            "output_file": str(self.output_path.name),
            "timestamp": datetime.datetime.now().isoformat(),
            "duration_seconds": duration,
            "sheets_processed": self.sheets_processed,
            "sheets_failed": self.sheets_failed,
            "error_message": self.error_message,
            "error_traceback": self.error_traceback,
        }

        # 2. Config Info
        if self.args:
            metadata["config_info"] = {
                "daf_mode": getattr(self.args, 'DAF', False),
                "custom_mode": getattr(self.args, 'custom', False),
                # Handle case where input_data_file might be a Path or string
                "input_file": str(Path(self.args.input_data_file).name) if hasattr(self.args, 'input_data_file') and self.args.input_data_file else "unknown",
                "config_dir": getattr(self.args, 'configdir', "./configs"),
                "generation_args": vars(self.args)
            }
        
        # 3. Database Export (Packing List Items)
        if self.input_data and "processed_tables_data" in self.input_data:
            processed_data = self.input_data["processed_tables_data"]
            packing_list_items = []
            
            # Iterate through tables (e.g., "1", "2")
            for table_id, table_data in processed_data.items():
                # All lists in table_data should be same length
                # We use 'po' or 'col_po' as the reference for row count
                row_count = 0
                possible_ref_keys = ["po", "col_po", "item", "col_item"]
                for k in possible_ref_keys:
                    if k in table_data and isinstance(table_data[k], list):
                        row_count = len(table_data[k])
                        break
                
                # Fallback: find any list
                if row_count == 0:
                    for v in table_data.values():
                        if isinstance(v, list):
                            row_count = len(v)
                            break
                
                for i in range(row_count):
                    try:
                        # Define strict mapping for database export fields
                        # keys = metadata output field
                        # values = list of possible source columns in the data (MUST start with col_)
                        field_map = {
                            "po": ["col_po"],
                            "item": ["col_item"],
                            "description": ["col_desc", "col_reference_code"],
                            "pcs": ["col_qty_pcs"],
                            "sqft": ["col_qty_sf"],
                            "pallet_count": ["col_pallet_count"],
                            "net": ["col_net"],
                            "gross": ["col_gross"],
                            "cbm": ["col_cbm"]
                        }

                        # Helper to safely get value at index using field map
                        def get_val(field_name, idx):
                            cols_to_try = field_map.get(field_name, [])
                            for col in cols_to_try:
                                lst = table_data.get(col)
                                if lst and isinstance(lst, list) and idx < len(lst):
                                    return lst[idx]
                            return None

                        item = {
                            "po": get_val("po", i),
                            "item": get_val("item", i),
                            "description": get_val("description", i),
                            "pcs": get_val("pcs", i),
                            "sqft": get_val("sqft", i),
                            "pallet_count": get_val("pallet_count", i),
                            "net": get_val("net", i),
                            "gross": get_val("gross", i),
                            "cbm": get_val("cbm", i)
                        }
                        packing_list_items.append(item)
                    except Exception:
                        continue # Skip malformed rows
            
            # 4. Summary Statistics
            # robust calculation handling strings/integers
            total_pcs = 0
            total_sqft = 0.0
            total_pallets = 0
            
            for i in packing_list_items:
                # PCS
                try: 
                    if i["pcs"] is not None: total_pcs += int(i["pcs"])
                except: pass
                
                # SQFT
                try:
                    if i["sqft"] is not None:
                        val_str = str(i["sqft"]).replace(',', '')
                        total_sqft += float(val_str)
                except: pass
                
                # PALLETS
                try:
                    if i["pallet_count"] is not None: total_pallets += int(i["pallet_count"])
                except: pass

            metadata["database_export"] = {
                "summary": {
                    "total_pcs": total_pcs,
                    "total_sqft": total_sqft,
                    "total_pallets": total_pallets,
                    "item_count": len(packing_list_items)
                },
                "packing_list_items": packing_list_items
            }
            
            # 7. Invoice Info (No, Date, Ref) - Extract from first table
            try:
                # Assuming data is in the first table "1" or first available
                first_table_key = next(iter(processed_data)) if processed_data else None
                if first_table_key:
                    table_data = processed_data[first_table_key]
                    
                    # Helper to safely get the first value from a list
                    def get_first(key):
                        val = table_data.get(key)
                        if isinstance(val, list) and val:
                            return val[0]
                        return None
    
                    metadata["invoice_info"] = {
                        "inv_no": get_first("col_inv_no"),
                        "inv_date": get_first("col_inv_date"),
                        "inv_ref": get_first("col_inv_ref")
                    }
            except Exception as e:
                metadata["invoice_info_error"] = str(e)

        # 5. Invoice Terms (from text replacements)
        if self.replacements_log:
            terms_found = set()
            for entry in self.replacements_log:
                if isinstance(entry, dict) and entry.get("term"):
                    terms_found.add(entry["term"])
            
            metadata["invoice_terms"] = {
                "detected_terms": list(terms_found),
                "replacements_detail": self.replacements_log
            }

        # 6. Header Info (Company Name, Address)
        if self.header_info:
            metadata["header_info"] = self.header_info
        
        # 7. Input Metadata (Legacy)
        metadata["input_metadata"] = self.input_data.get("metadata", {})

        try:
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=4)
            logger.info(f"Metadata written to {meta_path}")
        except Exception as e:
            logger.error(f"FATAL: Failed to write metadata: {e}")
