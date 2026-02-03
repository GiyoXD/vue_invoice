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


# Keep your existing imports
from core.invoice_generator.config.config_loader import BundledConfigLoader
from core.invoice_generator.builders.workbook_builder import WorkbookBuilder
from core.invoice_generator.builders.deep_sheet_builder import DeepSheetBuilder
from core.invoice_generator.processors.single_table_processor import SingleTableProcessor
from core.invoice_generator.processors.multi_table_processor import MultiTableProcessor
from core.invoice_generator.processors.placeholder_processor import PlaceholderProcessor

from core.invoice_generator.utils.print_area_config import configure_print_area
from core.invoice_generator.utils.generation_session import GenerationSession
from core.invoice_generator.resolvers import InvoiceAssetResolver

logger = logging.getLogger(__name__)



# --- Constants for Blueprints ---
from core.system_config import sys_config

DEFAULT_TEMPLATE_DIR = sys_config.templates_dir
DEFAULT_CONFIG_DIR = sys_config.registry_dir


# --- Helper Functions ---









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
    Uses GenerationSession context manager to ensure robust error handling.
    """
    # 1. Resolve Paths
    input_data_path, output_path, template_dir, config_dir = _resolve_generation_paths(
        input_data_path, output_path, template_dir, config_dir
    )

    # 2. Initialize Context
    ctx = _initialize_context(
        input_data_path, output_path, template_dir, config_dir,
        daf_mode, custom_mode, explicit_config_path, explicit_template_path, input_data_dict
    )

    # === CORE GENERATION LOGIC WITH MONITOR ===
    # Using 'meta_args' compatible dict for monitor (removing argparse dep)
    monitor_args = {
        "DAF": daf_mode,
        "custom": custom_mode,
        "input_data_file": str(input_data_path),
        "configdir": str(config_dir)
    }

    with GenerationSession(output_path, args=monitor_args, input_data=ctx.invoice_data) as session:
        logger.info("=== Starting Invoice Generation (Orchestrated) ===")
        
        # 3. Execution Pipeline
        _load_resources(ctx)
        monitor_paths = {
            'template': str(ctx.paths.get('template', 'unknown')),
            'config': str(ctx.paths.get('config', 'unknown'))
        }
        session.update_logs(header_info={"resolved_paths": monitor_paths})

        _prepare_workbooks(ctx)
        
        _process_sheets(ctx, session)
        
        _finalize(ctx)

    return output_path


# --- Internal Pipeline Structures ---

class GeneratorContext:
    """Holds state for the invoice generation pipeline."""
    def __init__(self, input_path: Path, output_path: Path, invoice_data: Dict):
        self.input_path = input_path
        self.output_path = output_path
        self.invoice_data = invoice_data
        
        # Paths
        self.template_dir: Optional[Path] = None
        self.config_dir: Optional[Path] = None
        self.paths: Dict[str, Path] = {}
        
        # Config & Resources
        self.config_loader: Optional[BundledConfigLoader] = None
        self.template_workbook: Optional[openpyxl.Workbook] = None
        self.output_workbook: Optional[openpyxl.Workbook] = None
        
        # Flags
        self.daf_mode = False
        self.custom_mode = False
        
        # Derived
        self.final_grand_total_pallets = 0


def _initialize_context(
    input_path: Path, output_path: Path, 
    template_dir: Path, config_dir: Path,
    daf_mode: bool, custom_mode: bool,
    manual_config: Optional[Path], manual_template: Optional[Path],
    data_dict: Optional[Dict]
) -> GeneratorContext:
    invoice_data = data_dict or {}
    if not invoice_data:
        logger.warning("No input data dictionary provided.")

    ctx = GeneratorContext(input_path, output_path, invoice_data)
    ctx.template_dir = template_dir
    ctx.config_dir = config_dir
    ctx.daf_mode = daf_mode
    ctx.custom_mode = custom_mode
    
    # Pre-resolve known manual paths
    if manual_config: ctx.paths['config'] = manual_config.resolve()
    if manual_template: ctx.paths['template'] = manual_template.resolve()
    
    return ctx


def _load_resources(ctx: GeneratorContext):
    """Stage 1: Resolve assets and load configuration."""
    # A. Resolve Paths
    resolver = InvoiceAssetResolver(base_config_dir=ctx.config_dir, base_template_dir=ctx.template_dir)
    assets = resolver.resolve_assets_for_input_file(str(ctx.input_path))
    
    if assets:
        if 'config' not in ctx.paths:
            ctx.paths['config'] = assets.config_path
            logger.info(f"Using resolved config: {ctx.paths['config']}")
        if 'template' not in ctx.paths:
            ctx.paths['template'] = assets.template_path
            logger.info(f"Using resolved template: {ctx.paths['template']}")
            
    ctx.paths['data'] = ctx.input_path

    # Validation
    if 'config' not in ctx.paths or 'template' not in ctx.paths:
         raise FileNotFoundError(f"Could not resolve config/template for '{ctx.input_path.name}'")

    # B. Load Config
    try:
        ctx.config_loader = BundledConfigLoader(ctx.paths['config'])
    except Exception as e:
        raise RuntimeError(f"Failed to load configuration: {e}") from e

    # C. Calculate Grand Totals (Legacy)
    processed_tables = ctx.invoice_data.get('processed_tables_data', {})
    if isinstance(processed_tables, dict):
        ctx.final_grand_total_pallets = sum(
            int(c) for t in processed_tables.values() 
            for c in (t.get("col_pallet_count") or t.get("pallet_count") or [])
            if str(c).isdigit()
        )


def _prepare_workbooks(ctx: GeneratorContext):
    """Stage 2: Load Template and Build Output Workbook."""
    ctx.output_path.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Loading template from: {ctx.paths['template']}")
    try:
        ctx.template_workbook = openpyxl.load_workbook(ctx.paths['template'], read_only=False)
    except Exception as e:
        # Fallback: JSON reconstruction
        json_config = ctx.config_loader.get_template_json_config()
        if json_config:
            logger.warning(f"Template load failed ({e}). Reconstructing from JSON.")
            ctx.template_workbook = openpyxl.Workbook()
            default_ws = ctx.template_workbook.active
            if default_ws: ctx.template_workbook.remove(default_ws)
            for sheet in json_config.keys():
                ctx.template_workbook.create_sheet(sheet)
        else:
            raise e

    # Build Output
    # We use the template workbook directly as the base.
    # This allows us to preserve "unknown sheets" (which are present in the XLSX but skipped in JSON)
    # intact in the final output. The Sanitizer has already cleaned the "Known Sheets".
    ctx.output_workbook = ctx.template_workbook

    # Deep Sheet Injection
    try:
        DeepSheetBuilder.build(ctx.output_workbook, ctx.invoice_data)
    except Exception as e:
        logger.warning(f"DeepSheet injection skipped: {e}")


def _process_sheets(ctx: GeneratorContext, session: GenerationSession):
    """Stage 3: Iterate and process each configured sheet."""
    sheets_config = ctx.config_loader.get_sheets_to_process()
    sheets_to_process = [s for s in sheets_config if s in ctx.output_workbook.sheetnames]
    
    if not sheets_to_process:
        raise ValueError("No valid sheets found to process.")

    # Mock args for processors (removing argparse dependency logic)
    # Processors expect an object with .DAF and .custom flags
    class ProcessorFlags:
        def __init__(self, daf, custom):
            self.DAF = daf
            self.custom = custom
    
    proc_args = ProcessorFlags(ctx.daf_mode, ctx.custom_mode)

    for sheet_name in sheets_to_process:
        logger.info(f"Processing sheet '{sheet_name}'")
        try:
            tmpl_ws = ctx.template_workbook[sheet_name]
            out_ws = ctx.output_workbook[sheet_name]
            sheet_conf = ctx.config_loader.get_sheet_config(sheet_name)
            ds_type = ctx.config_loader.get_data_source_type(sheet_name)
            
            if not ds_type:
                continue

            processor = _get_processor(
                ds_type, tmpl_ws, out_ws, sheet_name, sheet_conf, 
                ctx.config_loader, ctx.invoice_data, proc_args, ctx.final_grand_total_pallets,
                ctx.template_workbook, ctx.output_workbook 
            )

            if processor and processor.process():
                 session.log_success(sheet_name)
                 if hasattr(processor, 'replacements_log'):
                     session.update_logs(replacements=processor.replacements_log)
                 if hasattr(processor, 'header_info'):
                     session.update_logs(header_info=processor.header_info)
            else:
                 session.log_failure(sheet_name, error=RuntimeError("Processor returned False"))

        except Exception as e:
            session.log_failure(sheet_name, error=e)
            raise e


def _get_processor(ds_type, tmpl_ws, out_ws, name, conf, loader, data, args, pallets, tmpl_wb, out_wb):
    """Factory method for processors."""
    # Common kwargs
    kwargs = {
        "template_worksheet": tmpl_ws,
        "output_worksheet": out_ws,
        "sheet_name": name,
        "sheet_config": conf,
        "config_loader": loader,
        "data_source_indicator": ds_type,
        "invoice_data": data,
        "cli_args": args,
        "final_grand_total_pallets": pallets,
        "template_workbook": tmpl_wb,
        "output_workbook": out_wb
    }

    if ds_type in ["processed_tables_multi", "processed_tables"]:
        return MultiTableProcessor(**kwargs)
    elif ds_type == "placeholder":
        return PlaceholderProcessor(**kwargs)
    else:
        return SingleTableProcessor(**kwargs)


def _finalize(ctx: GeneratorContext):
    """Stage 4: Apply print settings and save."""
    logger.info("Applying Print Area & Page Setup...")
    for sheet in ctx.output_workbook.sheetnames:
        try:
            configure_print_area(ctx.output_workbook[sheet])
        except Exception as e:
            logger.error(f"Print setup failed for '{sheet}': {e}")

    logger.info(f"Saving workbook to {ctx.output_path}")
    ctx.output_workbook.save(ctx.output_path)
    
    # Cleanup
    if ctx.template_workbook: ctx.template_workbook.close()
    if ctx.output_workbook: ctx.output_workbook.close()


def _resolve_generation_paths(
    input_data_path: Path, 
    output_path: Path, 
    template_dir: Optional[Path] = None, 
    config_dir: Optional[Path] = None
) -> tuple[Path, Path, Path, Path]:
    """Resolves all paths and applies default directories if necessary."""
    # Ensure inputs are Path objects
    input_data_path = Path(input_data_path).resolve()
    output_path = Path(output_path).resolve()

    # Apply defaults for blueprint directories
    if template_dir is None:
        template_dir = DEFAULT_TEMPLATE_DIR
        logger.info(f"Using default blueprint template directory: {template_dir}")
    if config_dir is None:
        config_dir = DEFAULT_CONFIG_DIR
        logger.info(f"Using default blueprint config directory: {config_dir}")

    template_dir = Path(template_dir).resolve()
    config_dir = Path(config_dir).resolve()

    return input_data_path, output_path, template_dir, config_dir


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
    
    # Configure Logging for CLI using centralized logger
    from core.logger_config import setup_logging
    level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(log_dir=sys_config.run_log_dir, level=level)
    
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
        # Load data for CLI usage
        cli_data = {}
        try:
            with open(args.input_data_file, 'r', encoding='utf-8') as f:
                cli_data = json.load(f)
        except Exception as e:
            print(f"Failed to load input data file: {e}")
            sys.exit(1)

        run_invoice_generation(
            input_data_path=Path(args.input_data_file),
            output_path=output_path,
            template_dir=Path(args.templatedir) if args.templatedir else None,
            config_dir=Path(args.configdir) if args.configdir else None,
            daf_mode=args.DAF,
            custom_mode=args.custom,
            explicit_config_path=Path(args.config) if args.config else None,
            explicit_template_path=Path(args.template) if args.template else None,
            input_data_dict=cli_data
        )
        print(f"Successfully generated: {args.output}")
    except Exception as e:
        print(f"Generation failed: {e}")
        # The monitor would have already written the metadata file with the stack trace
        sys.exit(1)

if __name__ == "__main__":
    main()