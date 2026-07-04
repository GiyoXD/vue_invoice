# --- START OF FULL FILE: main.py ---
# --- Fixed datetime JSON serialization ---

import logging
import pprint
import decimal
import os
import json # Added for JSON output
import datetime # <<< ADDED IMPORT for datetime handling
from pathlib import Path # <<< ADDED IMPORT for pathlib
from typing import Dict, List, Any, Optional, Tuple, Union
import time # Added for timing operations

# --- Loop Profiler (non-invasive measurement) ---
from core.utils.loop_profiler import loop_profiler

# Import config directly
from . import config as cfg


from .excel_handler import ExcelHandler
from . import sheet_parser
from . import data_processor # Includes all processing functions
from .validation import DataValidationError, validate_data

# Use centralized logger - no basicConfig here
# Logging is configured by core.logger_config.setup_logging() at app startup
logger = logging.getLogger(__name__)

# --- Constants for Log Truncation ---
MAX_LOG_DICT_LEN = 3000 # Max length for printing large dicts in logs (for DEBUG)

# --- Constants for DAF Compounding Formatting ---
DAF_CHUNK_SIZE = 2  # How many items per group (e.g., PO1\\PO2)
DAF_INTRA_CHUNK_SEPARATOR = "/"  # Separator within a group (e.g., DOUBLE BACKSLASH)
DAF_INTER_CHUNK_SEPARATOR = "\n"  # Separator between groups (e.g., newline)



# --- >>> ADDED: Default JSON Serializer Function <<< ---
def json_serializer_default(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat() # Convert date/datetime to ISO string format
    elif isinstance(obj, decimal.Decimal): # Keep Decimal handling here too
        return float(obj)
    elif isinstance(obj, set): # Optional: Handle sets if needed
        return list(obj)
    # Add other custom types if needed
    # elif isinstance(obj, YourCustomClass):
    #     return obj.__dict__
    raise TypeError (f"Object of type {obj.__class__.__name__} is not JSON serializable")
# --- >>> END OF ADDED FUNCTION <<< ---


# Helper function to make data JSON serializable
# Handles tuple keys in aggregation results
def make_json_serializable(data):
    """Recursively converts tuple keys in dicts to strings and handles non-serializable types."""
    # NOTE: Using the default serializer for json.dumps handles Decimal and datetime now.
    # This function primarily focuses on converting tuple keys.
    if isinstance(data, dict):
        # Convert all keys to string, including tuple keys
        return {str(k): make_json_serializable(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [make_json_serializable(item) for item in data]
    elif data is None:
        return None # JSON null
    # Let the default handler in json.dumps deal with Decimal, datetime, etc.
    return data

# <<< MODIFIED FUNCTION SIGNATURE >>>
# Import PipelineMonitor

from core.utils.pipeline_monitor import PipelineMonitor
from core.utils.snitch import snitch

# ... (Previous code)

@snitch
def run_invoice_automation(
    input_excel_override: Union[str, Any] = None,
    input_filename_override: str = None,
    output_dir_override: str = None,
    ignore_tare_warning: bool = False,
    ignore_cbm_warning: bool = False
) -> Tuple[Path, str]:
    """
    Main entry point for the invoice automation process.
    Refactored to be callable as a library function.
    """
    # 1. Determine Output Directory (Fast Fail)
    if output_dir_override:
        output_dir = Path(output_dir_override).resolve()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise RuntimeError(f"Invalid output directory specified: {output_dir}")
    else:
        from core.system_config import sys_config
        output_dir = sys_config.temp_uploads_dir

    # 2. Determine Input File (Prep for Monitor)
    is_buffer = hasattr(input_excel_override, "read")
    if is_buffer:
        input_filepath = input_excel_override
        input_name = input_filename_override or "upload.xlsx"
    else:
        input_filepath = input_excel_override or getattr(cfg, 'INPUT_EXCEL_FILE', 'unknown.xlsx')
        input_name = Path(input_filepath).name
    
    # 3. Setup Monitor
    monitor_output_path = output_dir / f"{Path(input_name).stem}_parser.json"
    
    # WRAP EXECUTION
    with PipelineMonitor(monitor_output_path, step_name="Data Parser") as monitor:
        start_time = time.time()
        logging.info("--- Starting Invoice Automation ---")
        monitor.update_logs("input_file", input_name) # Record inputs

        # -------------------------------------------------------------
        # Re-Validate Input File inside Monitor to capture errors
        # -------------------------------------------------------------
        if not is_buffer:
            if not input_excel_override:
                 try:
                     input_filepath = cfg.INPUT_EXCEL_FILE
                     logging.info(f"Using input Excel path from config.py: {input_filepath}")
                 except Exception as e:
                     monitor.log_process_item("Configuration", status="error", error=e)
                     raise RuntimeError("Input Excel file path is missing in config.")

            if not os.path.isfile(input_filepath):
                 # Try relative resolution
                 script_dir = os.path.dirname(__file__)
                 potential_path = os.path.join(script_dir, input_filepath)
                 if os.path.isfile(potential_path):
                     input_filepath = potential_path
                     logging.info(f"Resolved relative input path: {input_filepath}")
                 else:
                     err = FileNotFoundError(f"Input Excel file not found: {input_filepath}")
                     monitor.log_process_item("Input File Check", status="error", error=err)
                     raise err
            input_filename = os.path.basename(input_filepath)
        else:
            input_filename = input_name

        logging.info(f"Processing workbook: {input_filename}")
        
        # ... [Rest of logic continues largely unchanged but inside this block] ...
        
        # CHANGED: processed_tables is now a list of tables, where each table is a list of row dicts
        processed_tables: List[List[Dict[str, Any]]] = []
        all_tables_data: List[List[Dict[str, Any]]] = []

        # Global definitions
        global_standard_aggregation_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]] = {}
        global_custom_aggregation_results: Dict[Tuple[Any, Any, Optional[str], None], Dict[str, decimal.Decimal]] = {}
        global_DAF_compounded_result: Optional[FinalDAFResultType] = None
        aggregation_mode_used = "standard"

        # Determine Aggregation Strategy
        use_custom_aggregation_for_DAF = False
        try:
            custom_prefixes = getattr(cfg, 'CUSTOM_AGGREGATION_WORKBOOK_PREFIXES', [])
            if not isinstance(custom_prefixes, list): custom_prefixes = []
            
            for prefix in custom_prefixes:
                 if input_filename.startswith(prefix):
                    use_custom_aggregation_for_DAF = True
                    aggregation_mode_used = "custom"
                    logging.info(f"Using CUSTOM aggregation (Prefix: {prefix})")
                    break
        except Exception as e:
            logging.error(f"Strategy determination error: {e}")
            monitor.log_warning(f"Aggregation strategy check failed: {e}")

        
        # --- PROCESSING STEPS ---
        try:
            logging.info(f"Loading workbook from: {input_filepath}")
            handler = ExcelHandler(input_filepath)
            sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True)
            if sheet is None: raise RuntimeError(f"Failed to load sheet from '{input_filepath}'.")
            
            actual_sheet_name = sheet.title
            
            # Header Detection
            smart_result = sheet_parser.find_and_map_smart_headers(sheet)
            if not smart_result: 
                 err = RuntimeError("Smart header detection failed.")
                 monitor.log_process_item("Header Detection", status="error", error=err)
                 raise err
            
            header_row, column_mapping = smart_result
            
            # --- Validate Required Columns ---
            # col_pallet_count is required for correct distribution of net, gross, and CBM values.
            # Without it, distribution bleeds across pallet boundaries producing silently wrong data.
            if 'col_pallet_count' not in column_mapping:
                err = DataValidationError(
                    "Required column 'col_pallet_count' was not found in the worksheet headers. "
                    "Please add a pallet count column (e.g. 'PALLET', '拖数', '件数', '托数') to the source Excel. "
                    "This column is required for correct distribution of net weight, gross weight, and CBM values."
                )
                monitor.log_process_item("Header Validation", status="error", error=err)
                raise err
            
            # Find Additional Tables
            additional_header_rows = sheet_parser.find_all_header_rows(
                sheet=sheet,
                search_pattern=cfg.HEADER_IDENTIFICATION_PATTERN,
                row_range=(header_row + 1, sheet.max_row),
                col_range=(cfg.HEADER_SEARCH_COL_RANGE[0], cfg.HEADER_SEARCH_COL_RANGE[1])
            )
            all_header_rows = [header_row] + additional_header_rows
            monitor.update_logs("tables_found", len(all_header_rows))
            
            all_tables_data = sheet_parser.extract_multiple_tables(sheet, all_header_rows, column_mapping)

            # Removed raw_tables_snapshot to avoid deepcopy overhead. Raw values are now saved in-place.
            # --- 5. Process Each Table (Instrumented) ---
            logging.info(f"--- Starting Data Processing Loop for {len(all_tables_data)} Extracted Table(s) ---")
            
            for table_index, current_table_data in enumerate(all_tables_data):
                table_id_str = f"Table {table_index + 1}"
                
                # Check for empty/invalid data (must be a non-empty list of dicts)
                if not isinstance(current_table_data, list) or not current_table_data:
                     monitor.log_warning(f"{table_id_str} is empty or invalid. Skipping.")
                     processed_tables.append([])
                     continue

                # --- 5.0.5: Normalize column types to Decimal/int ---
                data_processor.normalize_table_types(current_table_data)

                # --- 5.1: Validate Presence of Essential Data ---
                validate_data(current_table_data, table_id_str, column_mapping, monitor=monitor, phase='presence')
                
                try:
                    # 5a. CBM
                    data_after_cbm = data_processor.process_cbm_column(current_table_data)
                    
                    data_normalized = data_after_cbm

                    # 5b. Distribute
                    try:
                        # 5b.1 Strict Validation: Gross Weight MUST NOT be smaller than Net Weight (Before Distribution)
                        # This throws a DataValidationError if any single cell is inconsistent in the source data.
                        validate_data(data_normalized, table_id_str, column_mapping, monitor=monitor, phase='integrity', ignore_tare_warning=ignore_tare_warning)

                        data_after_distribution = data_processor.distribute_values(data_normalized, cfg.COLUMNS_TO_DISTRIBUTE, cfg.DISTRIBUTION_BASIS_COLUMN)
                        
                        # Validate distributed CBM and PCS proportions
                        validate_data(
                            data_after_distribution,
                            table_id_str,
                            column_mapping,
                            monitor=monitor,
                            phase='cbm_proportion',
                            ignore_cbm_warning=ignore_cbm_warning
                        )
                        
                        processed_tables.append(data_after_distribution)
                        data_for_aggregation = data_after_distribution
                    except DataValidationError as ve:
                        # Hard stop for validation errors (propagates to API)
                        raise ve
                    except Exception as distrib_e:
                        # Log but continue with fallback for non-critical distribution errors
                        monitor.log_warning(f"{table_id_str}: Distribution failed ({distrib_e}). Using raw/CBM data.")
                        processed_tables.append(data_after_cbm)
                        data_for_aggregation = data_after_cbm
                    
                    # 5c. Initial Aggregation
                    if data_for_aggregation:
                         data_processor.aggregate_standard_by_po_item_price(data_for_aggregation, global_standard_aggregation_results)
                         data_processor.aggregate_custom_by_po_item(data_for_aggregation, global_custom_aggregation_results)
                    
                    monitor.log_process_item(table_id_str, status="success")
                except DataValidationError as ve:
                    # User-facing validation errors MUST stop the whole process immediately.
                    # Reraise so it hits the outer catch-all and orchestrator.
                    raise ve
                except Exception as table_e:
                    # Log failure for this specific table but continue loop for general errors
                    monitor.log_process_item(table_id_str, status="error", error=table_e)
                    processed_tables.append([]) # Append empty to preserve indexing if needed


            # --- 6. DAF Compounding (Instrumented) ---
            try:
                # Determine strategy (re-using variables set earlier if accurate, or re-calculating simple version)
                # We reuse 'aggregation_mode_used' and 'use_custom_aggregation_for_DAF' calculated at start
                agg_source = global_custom_aggregation_results if use_custom_aggregation_for_DAF else global_standard_aggregation_results
                
                logging.info(f"Performing DAF Compounding (Mode: {aggregation_mode_used})")
                
                global_DAF_compounded_result = data_processor.perform_DAF_compounding(
                    processed_tables,
                    daf_chunk_size=DAF_CHUNK_SIZE,
                    daf_intra_separator=DAF_INTRA_CHUNK_SEPARATOR,
                    daf_inter_separator=DAF_INTER_CHUNK_SEPARATOR,
                )
                monitor.log_process_item("DAF Compounding", status="success")
            except Exception as daf_e:
                monitor.log_process_item("DAF Compounding", status="error", error=daf_e)

        except Exception as e:
            # Catch-all for the main Loading/Parsing/Extraction block
            monitor.log_process_item("Data Parsing/Extraction", status="error", error=e)
            raise e # Re-raise to trigger outer exit codes if needed, though monitor captures it



        # --- 7. Output / Further Steps ---
        logging.info(f"Final processed data structure contains {len(processed_tables)} table(s).")
        logging.info(f"Primary aggregation mode used for DAF Compounding: {aggregation_mode_used.upper()}")




        # --- Calculate Add-on Data (Leather Summary) ---
        logging.info("--- Calculating Add-on Data ---")
        
        # Calculate grand total (merged across all tables)
        merged_processed_data: List[Dict[str, Any]] = []
        for table_data in processed_tables:
            if isinstance(table_data, list):
                merged_processed_data.extend(table_data)

        # Calculate normal aggregate per PO with pallets (group by PO + price)
        # We calculate this FIRST so we can use its accurate pallet count for the footer
        normal_aggregate_per_po = data_processor.aggregate_per_po_with_pallets(merged_processed_data)
        logging.info(f"Normal Aggregate Per PO: {len(normal_aggregate_per_po)} unique PO+price combinations")

        # Extract the true total pallets from the aggregated manifest
        true_total_pallets = sum(item.get('col_pallet_count', 0) for item in normal_aggregate_per_po)

        # Calculate leather summary (BUFFALO vs COW) across all tables
        # Use the normal_aggregate_per_po data so we get integer pallet counts and correct sums
        leather_summary = data_processor.calculate_leather_summary(normal_aggregate_per_po)
        logging.info(f"Leather Summary: {leather_summary}")

        # Calculate weight summary across all tables
        raw_weight_summary = data_processor.calculate_weight_summary(merged_processed_data)
        weight_summary_addon = {
            'net': float(raw_weight_summary.get('col_net', 0.0)),
            'gross': float(raw_weight_summary.get('col_gross', 0.0))
        }
        logging.info(f"Weight Summary Addon: {weight_summary_addon}")

        # --- Calculate Footer Data ---
        logging.info("--- Calculating Footer Data ---")
        
        # Calculate per-table totals
        table_footer_data = []
        for table_index, table_data in enumerate(processed_tables):
            table_id = str(table_index + 1)
            if isinstance(table_data, list):
                footer_totals = data_processor.calculate_footer_totals(table_data)
                
                # If there's only one table, it gets all the pallets. If multiple, we'd need to distribute, 
                # but for now we'll rely on the parser to not double count. (Will be fixed in data_processor.py)
                
                table_footer_data.append(footer_totals)
                logging.info(f"Table {table_id} Footer: {footer_totals}")
        
        # Calculate grand total (merged across all tables)
        grand_total_footer = data_processor.calculate_footer_totals(merged_processed_data)
        
        # Override the potentially inflated pallet count with the true aggregated count
        grand_total_footer['col_pallet_count'] = true_total_pallets
        
        # If there's only one table (common case), ensure its table total also has the correct pallet count
        if len(table_footer_data) == 1:
            table_footer_data[0]['col_pallet_count'] = true_total_pallets

        logging.info(f"Grand Total Footer: {grand_total_footer}")

        # Format pallet counts to "x-y" display format in-place for JSON output
        data_processor.format_pallet_counts_to_xy(processed_tables, true_total_pallets)
        data_processor.format_pallet_counts_to_xy([normal_aggregate_per_po], true_total_pallets)
        # Remove col_pallet_id from final output structures as it is strictly for validation
        for table in processed_tables:
            for row in table:
                row.pop('col_pallet_id', None)
        # --- 8. Generate JSON Output ---
        logging.info("--- Preparing Data for JSON Output ---")
        try:
            # Create the structure to be converted to JSON
            # Use the helper function to ensure serializability
            final_json_structure = {
                 "metadata": {
                    "workbook_filename": input_filename, # Use the actual input filename
                    "worksheet_name": actual_sheet_name,
                    "DAF_compounding_input_mode": aggregation_mode_used, # Clarify which mode fed DAF
                    "DAF_chunk_size": DAF_CHUNK_SIZE,
                     "DAF_intra_separator": DAF_INTRA_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'), # Encode escapes for JSON clarity
                    "DAF_inter_separator": DAF_INTER_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'), # Encode escapes for JSON clarity
                    "timestamp": datetime.datetime.now().isoformat(), # Add generation timestamp
                    "warnings": monitor.warnings # Surface runtime warnings to frontend
                },
                "price_adjustment": [], # Initialized for frontend adjustments
                 # Include processed table data (potentially large)
                 # RENAME: processed_tables_data -> multi_table
                 "multi_table": make_json_serializable(processed_tables),

                 # Raw/unprocessed table data exactly as extracted from Excel.
                 # CBM and other values are NEVER distributed here.
                  # Kept purely for backward compatibility with old frontend/db queries.
                  "raw_data": [],
                 
                 # Include Footer Data - both per-table and grand total
                 "footer_data": {
                     "table_totals": make_json_serializable(table_footer_data),  # Per-table totals
                     "grand_total": make_json_serializable(grand_total_footer),   # Overall grand total
                     "add_ons": {
                         "leather_summary_addon": make_json_serializable(leather_summary),  # BUFFALO vs COW summary
                         "weight_summary_addon": make_json_serializable(weight_summary_addon),
                     }
                 },

                # Group all unified aggregation outputs under single_table
                "single_table": {
                    # Include BOTH aggregation results explicitly (formatted as lists)
                    # RENAME: standard_aggregation_results -> aggregation (Matches Config)
                    "aggregation": data_processor.format_aggregation_as_list(global_standard_aggregation_results, mode='standard'),
                    # RENAME: custom_aggregation_results -> aggregation_custom (Matches Suffix Rule)
                    "aggregation_custom": data_processor.format_aggregation_as_list(global_custom_aggregation_results, mode='custom'),
                    
                    # Normal aggregate per PO with pallets (group by PO + price)
                    # RENAME: normal_aggregate_per_po_with_pallets -> manifest_by_pallet_per_po (User Request)
                    "manifest_by_pallet_per_po": make_json_serializable(normal_aggregate_per_po),

                    # Include the final compounded result (derived from one of the above, based on mode)
                    # RENAME: final_DAF_compounded_result -> aggregation_DAF (Matches Suffix Rule)
                    "aggregation_DAF": make_json_serializable(global_DAF_compounded_result)
                }
            }

             # Convert the structure to a JSON string (pretty-printed)
            json_output_string = json.dumps(final_json_structure,
                                            indent=4,
                                            default=json_serializer_default) # Use the default serializer

            # Do not log raw JSON output to keep console output clean
            logging.info(f"Generated JSON output structure successfully ({len(json_output_string)} chars).")

            # --- MODIFIED: Save JSON using output_dir and simplified filename ---
            input_stem = Path(input_filename).stem # Get filename without extension
            json_output_filename = f"{input_stem}.json" # Simplified filename
            output_json_path = output_dir / json_output_filename # Combine output dir and filename

            logging.info(f"Determined output JSON path: {output_json_path}")
            try:
                # --- Atomic Write: write to temp file, verify, then rename ---
                # This prevents truncated/corrupt JSON from being visible to consumers.
                import tempfile
                temp_fd, temp_path = tempfile.mkstemp(
                    suffix='.json.tmp', dir=str(output_json_path.parent)
                )
                try:
                    with os.fdopen(temp_fd, 'w', encoding='utf-8') as f_json:
                        f_json.write(json_output_string)
                        f_json.flush()
                        os.fsync(f_json.fileno())  # Force write to disk
                    
                    # Post-write integrity check: read back and parse to verify
                    with open(temp_path, 'r', encoding='utf-8') as f_verify:
                        json.load(f_verify)  # Will raise JSONDecodeError if truncated/corrupt
                    
                    # Verification passed — atomically replace the target file
                    import shutil
                    shutil.move(temp_path, str(output_json_path))
                    logging.info(f"Successfully saved JSON output to '{output_json_path}' (verified)")
                except Exception:
                    # Clean up temp file on any failure
                    if os.path.exists(temp_path):
                        os.unlink(temp_path)
                    raise
            except json.JSONDecodeError as verify_err:
                logging.error(f"CRITICAL: JSON integrity check failed after write — output would be truncated/corrupt: {verify_err}")
                raise RuntimeError(
                    f"JSON output verification failed: the generated data could not be re-parsed. "
                    f"This usually means the data is too large or contains unserializable values. "
                    f"Details: {verify_err}"
                )
            except IOError as io_err:
                logging.error(f"Failed to write JSON output to file '{output_json_path}': {io_err}")
                raise io_err
            except Exception as write_err:
                 logging.error(f"An unexpected error occurred while writing JSON file: {write_err}", exc_info=True)
                 raise write_err

        except TypeError as json_err:
            logging.error(f"Failed to serialize data to JSON: {json_err}. Check data types and default handler.", exc_info=True)
            raise json_err
        except Exception as e:
            logging.error(f"An unexpected error occurred during JSON generation: {e}", exc_info=True)
            raise e
        # --- End JSON Generation ---

        logging.info(f"📁 Processed file: {input_filename}")

        # --- Profiler Report (non-invasive) ---
        loop_profiler.report(title=f"Sheet Parser Profiler — {input_filename}")
        loop_profiler.reset()
        
        return output_json_path, input_stem