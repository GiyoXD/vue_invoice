# core/invoice_generator/generate_invoice.py
import os
import json
import pickle
import argparse
import sys
import time
import datetime
import logging
import traceback
from pathlib import Path
from typing import Optional, Dict, Any, List
import openpyxl
import re
import ast

if __name__ == "__main__":
    import sys
    from pathlib import Path
    try:
        # Resolve project root (assuming core/invoice_generator/generate_invoice.py structure)
        root_path = Path(__file__).resolve().parents[2]
        if str(root_path) not in sys.path:
            sys.path.insert(0, str(root_path))
    except IndexError:
        pass

# Keep your existing imports
from core.invoice_generator.config.config_loader import BundledConfigLoader
from core.invoice_generator.builders.workbook_builder import WorkbookBuilder
from core.invoice_generator.processors.single_table_processor import SingleTableProcessor
from core.invoice_generator.processors.multi_table_processor import MultiTableProcessor
from core.invoice_generator.processors.placeholder_processor import PlaceholderProcessor
from core.invoice_generator.utils.monitor import GenerationMonitor
from core.invoice_generator.resolvers import InvoiceAssetResolver

logger = logging.getLogger(__name__)

# --- Global Exception Hook ---
def global_exception_handler(exc_type, exc_value, exc_traceback):
    """
    Catch uncaught exceptions (e.g., SyntaxError, ImportError) that happen
    before the GenerationMonitor context is entered.
    """
    logger.critical("Uncaught exception (Pre-Monitor)", exc_info=(exc_type, exc_value, exc_traceback))
    # Optionally write a "crash.json" here if needed, but logger is usually enough for pre-start crashes

sys.excepthook = global_exception_handler

# --- Constants for Blueprints ---
from core.system_config import sys_config

DEFAULT_TEMPLATE_DIR = sys_config.templates_dir
DEFAULT_CONFIG_DIR = sys_config.registry_dir


# --- Helper Functions ---
def derive_paths(input_data_path: str, template_dir: str, config_dir: str) -> Optional[Dict[str, Path]]:
    """
    Derive paths for config and template based on input data filename.
    """
    input_path = Path(input_data_path)
    stem = input_path.stem
    
    # Prioritize bundle config to avoid picking up data file as config
    config_path = Path(config_dir) / f"{stem}_bundle_config.json"
    
    # Heuristic: If exact match not found, try stripping trailing numbers/underscores (e.g., JF25057 -> JF)
    effective_stem = stem
    if not config_path.exists():
        prefix = re.sub(r'[\d_]+$', '', stem)
        if prefix and prefix != stem:
            prefix_config = Path(config_dir) / f"{prefix}_bundle_config.json"
            if prefix_config.exists():
                config_path = prefix_config
                effective_stem = prefix # Use the prefix for template lookup too
                logger.info(f"Found config using prefix match: '{stem}' -> '{prefix}'")

    if not config_path.exists():
        config_path = Path(config_dir) / f"{stem}.json"
    
    # Fallback to default config if specific not found
    if not config_path.exists():
        default_config = Path(config_dir) / "default.json"
        if default_config.exists():
             config_path = default_config
        else:
             # If no config found, we can't proceed unless we have a strategy
             pass

    # Template path - ideally derived from config, but we need config first.
    # Strategy: Load config, check for template name. If not, use stem.
    template_path = None
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                # Check if template is specified in meta or processing
                template_name = cfg.get('_meta', {}).get('template_name')
                if template_name:
                    template_path = Path(template_dir) / template_name
        except:
            pass
    
    if not template_path:
        # Try effective stem first (e.g. JF.xlsx)
        template_path = Path(template_dir) / f"{effective_stem}.xlsx"
        if not template_path.exists() and effective_stem != stem:
             # Try original stem if effective stem failed (e.g. JF25057.xlsx)
             template_path = Path(template_dir) / f"{stem}.xlsx"

        if not template_path.exists():
             # Fallback to generic Invoice.xlsx or configured default (JF.xlsx)
             fallback = Path(template_dir) / sys_config.default_template_name
             if fallback.exists():
                 template_path = fallback

    if config_path.exists() and template_path and template_path.exists():
        return {
            'data': input_path,
            'config': config_path,
            'template': template_path
        }
    
    logger.error(f"Could not derive paths. Config: {config_path} (Exists: {config_path.exists()}), Template: {template_path} (Exists: {template_path and template_path.exists()})")
    return None

def load_data(path: Path) -> Dict[str, Any]:
    """Load invoice data from JSON."""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Failed to load data from {path}: {e}")
        return {}

def run_invoice_generation(
    input_data_path: Path,
    output_path: Path,
    template_dir: Optional[Path] = None,
    config_dir: Optional[Path] = None,
    daf_mode: bool = False,
    custom_mode: bool = False,
    explicit_config_path: Optional[Path] = None,
    explicit_template_path: Optional[Path] = None,
    input_data_dict: Optional[Dict[str, Any]] = None
) -> Path:
    """
    Library entry point for invoice generation. 
    Uses GenerationMonitor context manager to ensure robust error handling and metadata generation.
    """
    # Ensure paths are Path objects
    input_data_path = Path(input_data_path).resolve()
    output_path = Path(output_path).resolve()

    # Use defaults if not provided
    if template_dir is None:
        template_dir = DEFAULT_TEMPLATE_DIR
        logger.info(f"Using default blueprint template directory: {template_dir}")
    if config_dir is None:
        config_dir = DEFAULT_CONFIG_DIR
        logger.info(f"Using default blueprint config directory: {config_dir}")

    template_dir = Path(template_dir).resolve()
    config_dir = Path(config_dir).resolve()

    # Mock CLI args for metadata generation
    meta_args = argparse.Namespace(
        DAF=daf_mode, 
        custom=custom_mode, 
        input_data_file=str(input_data_path), 
        configdir=str(config_dir)
    )

    # Pre-load data to pass to monitor (for input metadata)
    invoice_data = {}
    if input_data_dict:
        invoice_data = input_data_dict
    else:
        # Attempt minimal load for metadata if possible, otherwise monitor will handle empty
        try:
             with open(input_data_path, 'r', encoding='utf-8') as f:
                invoice_data = json.load(f)
        except:
            pass

    # === CORE GENERATION LOGIC WITH MONITOR ===
    with GenerationMonitor(output_path, args=meta_args, input_data=invoice_data) as monitor:
        
        logger.info("=== Starting Invoice Generation (Library Call) ===")
        logger.debug(f"Input: {input_data_path}, Output: {output_path}")

        # 1. Derive Paths (Using Resolver)
        resolver = InvoiceAssetResolver(base_config_dir=config_dir, base_template_dir=template_dir)
        assets = resolver.resolve_assets_for_input_file(str(input_data_path))
        
        # Determine actual paths to use
        paths = {}
        
        # Priority: Explicit -> Resolved -> Error
        if explicit_config_path:
            paths['config'] = explicit_config_path.resolve()
            logger.info(f"Using explicit config path: {paths['config']}")
        elif assets:
            paths['config'] = assets.config_path
            logger.info(f"Using resolved config path: {paths['config']}")
            
        if explicit_template_path:
            paths['template'] = explicit_template_path.resolve()
            logger.info(f"Using explicit template path: {paths['template']}")
        elif assets:
            paths['template'] = assets.template_path
            logger.info(f"Using resolved template path: {paths['template']}")

        paths['data'] = input_data_path

        # Validate we exist
        if 'config' not in paths or 'template' not in paths:
             # Try fallback to legacy just in case resolver completely failed or weird explicit combo
             # But resolver handles fallback.
             if not assets and not (explicit_config_path or explicit_template_path):
                 raise FileNotFoundError(f"Could not derive template/config paths for {input_data_path.name}")
             else:
                 if 'config' not in paths: raise FileNotFoundError("Missing config path")
                 if 'template' not in paths: raise FileNotFoundError("Missing template path")

        # 2. Load Configuration
        try:
            config_loader = BundledConfigLoader(paths['config'])
        except Exception as e:
            raise RuntimeError(f"Failed to load configuration: {e}") from e
        
        # 3. Load/Verify Data (Already loaded for monitor, but let's ensure it's valid for processing)
        if not invoice_data:
             # Retry load properly using the helper which logs errors
             invoice_data = load_data(paths['data'])
             
        if not invoice_data:
            raise RuntimeError("No invoice data available or failed to load data")

        # 4. Prepare Output Directory
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        template_workbook = None
        output_workbook = None
        
        try:
            # Step 5: Load Template & Create Workbook
            logger.info(f"Loading template from: {paths['template']}")
            try:
                template_workbook = openpyxl.load_workbook(paths['template'], read_only=False)
            except Exception as e:
                # Fallback: Check if we have JSON config to reconstruct from scratch
                json_config = config_loader.get_template_json_config()
                if json_config:
                    logger.warning(f"Failed to load Excel template ({e}). Found template JSON, proceeding with reconstruction.")
                    template_workbook = openpyxl.Workbook()
                    # Remove default sheet
                    default_ws = template_workbook.active
                    if default_ws:
                        template_workbook.remove(default_ws)
                    
                    # Create empty sheets based on JSON keys (Processors don't need content in template_worksheet for JSON path)
                    for sheet_name in json_config.keys():
                        template_workbook.create_sheet(sheet_name)
                    logger.info(f"Created dummy template workbook with sheets: {template_workbook.sheetnames}")
                else:
                    logger.error(f"Failed to load template and no sibling JSON config found: {e}")
                    raise e
            
            workbook_builder = WorkbookBuilder(sheet_names=template_workbook.sheetnames)
            output_workbook = workbook_builder.build()
            
            # Step 6: Determine Sheets to Process
            sheets_to_process_config = config_loader.get_sheets_to_process()
            sheets_to_process = [s for s in sheets_to_process_config if s in output_workbook.sheetnames]

            if not sheets_to_process:
                logger.error(f"DEBUG: Config Sheets: {sheets_to_process_config}")
                logger.error(f"DEBUG: Workbook Sheets: {list(output_workbook.sheetnames) if output_workbook else 'None'}")
                raise ValueError("No valid sheets found to process in configuration.")

            # Global calculation (legacy support)
            final_grand_total_pallets = 0
            processed_tables = invoice_data.get('processed_tables_data', {})
            if isinstance(processed_tables, dict):
                final_grand_total_pallets = sum(
                    int(c) for t in processed_tables.values() 
                    for c in t.get("pallet_count", []) 
                    if str(c).isdigit()
                )

            # Step 7: Processing Loop
            for sheet_name in sheets_to_process:
                logger.info(f"Processing sheet '{sheet_name}'")
                
                # Try/Except per sheet to allow partial success if desired (or just fail fast)
                try: 
                    template_worksheet = template_workbook[sheet_name]
                    output_worksheet = output_workbook[sheet_name]
                    
                    sheet_config = config_loader.get_sheet_config(sheet_name)
                    data_source_indicator = config_loader.get_data_source_type(sheet_name)

                    if not data_source_indicator:
                        logger.warning(f"Skipping '{sheet_name}': No data source configured.")
                        continue

                    # Instantiate Processor
                    mock_args = argparse.Namespace(DAF=daf_mode, custom=custom_mode)

                    processor = None
                    if data_source_indicator in ["processed_tables_multi", "processed_tables"]:
                        processor = MultiTableProcessor(
                            template_workbook=template_workbook,
                            output_workbook=output_workbook,
                            template_worksheet=template_worksheet,
                            output_worksheet=output_worksheet,
                            sheet_name=sheet_name,
                            sheet_config=sheet_config,
                            config_loader=config_loader,
                            data_source_indicator=data_source_indicator,
                            invoice_data=invoice_data,
                            cli_args=mock_args, 
                            final_grand_total_pallets=final_grand_total_pallets
                        )
                    elif data_source_indicator == "placeholder":
                        processor = PlaceholderProcessor(
                            template_workbook=template_workbook,
                            output_workbook=output_workbook,
                            template_worksheet=template_worksheet,
                            output_worksheet=output_worksheet,
                            sheet_name=sheet_name,
                            sheet_config=sheet_config,
                            config_loader=config_loader,
                            data_source_indicator=data_source_indicator,
                            invoice_data=invoice_data,
                            cli_args=mock_args, 
                            final_grand_total_pallets=final_grand_total_pallets
                        )
                    else:
                        processor = SingleTableProcessor(
                            template_workbook=template_workbook,
                            output_workbook=output_workbook,
                            template_worksheet=template_worksheet,
                            output_worksheet=output_worksheet,
                            sheet_name=sheet_name,
                            sheet_config=sheet_config,
                            config_loader=config_loader,
                            data_source_indicator=data_source_indicator,
                            invoice_data=invoice_data,
                            cli_args=mock_args, 
                            final_grand_total_pallets=final_grand_total_pallets
                        )

                    if processor:
                        if processor.process():
                            monitor.log_success(sheet_name)
                            # Collect logs
                            if hasattr(processor, 'replacements_log'):
                                monitor.update_logs(replacements=processor.replacements_log)
                            if hasattr(processor, 'header_info'):
                                monitor.update_logs(header_info=processor.header_info)
                        else:
                            monitor.log_failure(sheet_name, error=RuntimeError("Processor returned False"))
                
                except Exception as e:
                    monitor.log_failure(sheet_name, error=e)
                    # We continue to next sheet? Or raise?
                    # Generally for invoices, if one sheet blocks, it might be safer to fail?
                    # But the monitor allows us to choose. Let's re-raise to be safe for now 
                    # unless we want "partial_success".
                    raise e 

            # Step 8: Save
            logger.info(f"Saving workbook to {output_path}")
            output_workbook.save(output_path)
            
        finally:
            if template_workbook: template_workbook.close()
            if output_workbook: output_workbook.close()

    return output_path

def main():
    """CLI Entry point for backward compatibility."""
    parser = argparse.ArgumentParser(description="Generate Invoice CLI")
    parser.add_argument("input_data_file", help="Path to input data file")
    parser.add_argument("-o", "--output", default=None, help="Output path (default: output/ dir)")
    parser.add_argument("-t", "--templatedir", default=None, help="Template dir (defaults to database/blueprints/template)")
    parser.add_argument("-c", "--configdir", default=None, help="Config dir (defaults to database/blueprints/config/bundled)")
    parser.add_argument("--config", help="Explicit path to config file")
    parser.add_argument("--template", help="Explicit path to template file")
    parser.add_argument("--DAF", action="store_true", help="DAF mode")
    parser.add_argument("--custom", action="store_true", help="Custom mode")
    parser.add_argument("--debug", action="store_true", help="Debug logging")
    
    args = parser.parse_args()
    
    # Configure Logging for CLI
    level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(level=level, format='%(levelname)s: %(message)s')
    
    # Determine Output Path
    if args.output:
        output_path = Path(args.output)
    else:
        from core.system_config import sys_config
        # Derive from input stem
        input_stem = Path(args.input_data_file).stem
        output_path = sys_config.output_dir / f"{input_stem}.xlsx"
        # Ensure output dir exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        run_invoice_generation(
            input_data_path=Path(args.input_data_file),
            output_path=output_path,
            template_dir=Path(args.templatedir) if args.templatedir else None,
            config_dir=Path(args.configdir) if args.configdir else None,
            daf_mode=args.DAF,
            custom_mode=args.custom,
            explicit_config_path=Path(args.config) if args.config else None,
            explicit_template_path=Path(args.template) if args.template else None
        )
        print(f"Successfully generated: {args.output}")
    except Exception as e:
        print(f"Generation failed: {e}")
        # The monitor would have already written the metadata file with the stack trace
        sys.exit(1)

if __name__ == "__main__":
    main()