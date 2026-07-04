# --- START OF FULL FILE: main.py ---
# --- Fixed datetime JSON serialization ---

import logging
import pprint
import decimal
import os
from pathlib import Path
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
from .util.exporter import export_invoice_data

# Use centralized logger - no basicConfig here
# Logging is configured by core.logger_config.setup_logging() at app startup
logger = logging.getLogger(__name__)

# --- Constants for Log Truncation ---
MAX_LOG_DICT_LEN = 3000 # Max length for printing large dicts in logs (for DEBUG)



# --- # (Serialization functions refactored and moved to core.data_parser.util.serializer.Serializer)

# <<< MODIFIED FUNCTION SIGNATURE >>>
# Import PipelineMonitor

from core.utils.pipeline_monitor import PipelineMonitor
from core.utils.snitch import snitch

# ... (Previous code)

@snitch
def main(
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
                    daf_chunk_size=cfg.DAF_CHUNK_SIZE,
                    daf_intra_separator=cfg.DAF_INTRA_CHUNK_SEPARATOR,
                    daf_inter_separator=cfg.DAF_INTER_CHUNK_SEPARATOR,
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
        logging.info(f"Grand Total Footer: {grand_total_footer}")


        # --- 8. Export Results to JSON ---
        input_stem = Path(input_filename).stem
        output_json_path = export_invoice_data(
            output_dir=output_dir,
            input_filename=input_filename,
            input_stem=input_stem,
            actual_sheet_name=actual_sheet_name,
            aggregation_mode_used=aggregation_mode_used,
            processed_tables=processed_tables,
            table_footer_data=table_footer_data,
            grand_total_footer=grand_total_footer,
            leather_summary=leather_summary,
            weight_summary_addon=weight_summary_addon,
            global_standard_aggregation_results=global_standard_aggregation_results,
            global_custom_aggregation_results=global_custom_aggregation_results,
            normal_aggregate_per_po=normal_aggregate_per_po,
            global_DAF_compounded_result=global_DAF_compounded_result,
            warnings=monitor.warnings
        )

        logging.info(f"📁 Processed file: {input_filename}")

        # --- Profiler Report (non-invasive) ---
        loop_profiler.report(title=f"Sheet Parser Profiler — {input_filename}")
        loop_profiler.reset()
        
        return output_json_path, input_stem