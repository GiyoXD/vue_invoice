# core/invoice_generator/generate_invoice.py
import json
import argparse
import sys
import io
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Dict, Any, List
import openpyxl

from core.invoice_generator.config import ConfigFileReader, ConfigStore
from core.invoice_generator.builders.deep_sheet_builder import DeepSheetBuilder
from core.invoice_generator.processors.single_table_processor import SingleTableProcessor
from core.invoice_generator.processors.multi_table_processor import MultiTableProcessor
from core.invoice_generator.models.context import (
    ProcessorContext, ExcelIOContext, SheetConfigContext, RuntimeDataContext
)
from core.invoice_generator.models.request import (
    GenerationOptions, InvoicePathConfig, ExplicitOverrides, InvoiceGenerationRequest
)

from core.invoice_generator.utils.generation_session import GenerationSession
from core.invoice_generator.utils.workbook_utils import (
    inject_unknown_sheets,
    apply_print_settings,
    build_output_filename,
    finalize,
    split_workbook_to_buffers,
)
from core.invoice_generator.resolvers import InvoiceAssetResolver
from core.system_config import sys_config

logger = logging.getLogger(__name__)

# --- Constants for Blueprints ---
DEFAULT_TEMPLATE_DIR = sys_config.templates_dir
DEFAULT_CONFIG_DIR = sys_config.registry_dir


def run_invoice_generation(req: InvoiceGenerationRequest):
    """
    Library entry point for invoice generation. 
    Uses GenerationSession context manager to ensure robust error handling.
    """
    opts = req.options or GenerationOptions()

    # 0. FORCE CLEAR SESSION LOG (Per User Request)
    try:
        from core.logger_config import clear_session_log
        clear_session_log()
    except Exception:
        pass

    # 1. Resolve Paths
    input_data_path, output_path, template_dir, config_dir = _resolve_generation_paths(
        req.paths.input_data_path, req.paths.output_path, req.paths.template_dir, req.paths.config_dir
    )

    # 2. Initialize Context
    ctx = _initialize_context(
        input_data_path, output_path, template_dir, config_dir, req
    )

    # === CORE GENERATION LOGIC WITH MONITOR ===
    # Using 'meta_args' compatible dict for monitor (removing argparse dep)
    monitor_args = {
        "DAF": opts.daf_mode,
        "custom": opts.custom_mode,
        "input_data_file": str(input_data_path),
        "configdir": str(config_dir)
    }

    with GenerationSession(output_path, args=monitor_args, input_data=ctx.invoice_data) as session:
        logger.info("=== Starting Invoice Generation (Orchestrated) ===")
        
        # 3. Execution Pipeline
        try:
            _load_resources(ctx)
        except ValueError as e:
            logger.error(f"[_load_resources] Failed to load resources: {e}")
            raise
        monitor_paths = {
            'template': str(ctx.paths.get('template', 'unknown')),
            'config': str(ctx.paths.get('config', 'unknown'))
        }
        session.update_logs(header_info={"resolved_paths": monitor_paths})

        _prepare_workbooks(ctx)
        
        _process_sheets(ctx, session)
        
        inject_unknown_sheets(ctx)
        
    # 4. Build dynamic output filename based on sheets & invoice_id
        build_output_filename(ctx)
        
        if opts.return_bytes:
            apply_print_settings(ctx)
            
            if opts.split_sheets:
                output_files = split_workbook_to_buffers(ctx.output_workbook, ctx.output_path)
                if ctx.output_workbook: ctx.output_workbook.close()
                return output_files
            else:
                logger.info("Saving workbook to in-memory buffer")
                buffer = io.BytesIO()
                ctx.output_workbook.save(buffer)
                if ctx.output_workbook: ctx.output_workbook.close()
                return ctx.output_path.name, buffer.getvalue()
        else:
            finalize(ctx)

    return ctx.output_path


# --- Internal Pipeline Structures ---

class GeneratorContext:
    """Holds state for the invoice generation pipeline."""
    def __init__(self, input_path: Path, output_path: Path, invoice_data: Dict,
                 options: Optional[GenerationOptions] = None):
        self.input_path = input_path
        self.output_path = output_path
        self.invoice_data = invoice_data
        self.options = options or GenerationOptions()
        
        # Paths
        self.template_dir: Optional[Path] = None
        self.config_dir: Optional[Path] = None
        self.paths: Dict[str, Path] = {}
        
        # Config & Resources
        self.config_loader: Optional[ConfigStore] = None
        self.template_workbook: Optional[openpyxl.Workbook] = None
        self.output_workbook: Optional[openpyxl.Workbook] = None
        
        # Derived
        self.template_xlsx_bytes: Optional[bytes] = None

    # --- Convenience accessors for backward compatibility ---
    @property
    def daf_mode(self): return self.options.daf_mode
    @property
    def custom_mode(self): return self.options.custom_mode
    @property
    def enable_auto_fit(self): return self.options.enable_auto_fit
    @property
    def split_sheets(self): return self.options.split_sheets


def _initialize_context(
    input_path: Path, output_path: Path, 
    template_dir: Path, config_dir: Path,
    req: InvoiceGenerationRequest
) -> GeneratorContext:
    invoice_data = req.overrides.input_data_dict or {}
    if not invoice_data:
        logger.warning("No input data dictionary provided.")

    ctx = GeneratorContext(input_path, output_path, invoice_data, req.options)
    ctx.template_dir = template_dir
    ctx.config_dir = config_dir
    
    # Pre-resolve known manual paths
    if req.overrides.explicit_config_path: ctx.paths['config'] = req.overrides.explicit_config_path.resolve()
    if req.overrides.explicit_template_path: ctx.paths['template'] = req.overrides.explicit_template_path.resolve()
    
    return ctx


def _load_resources(ctx: GeneratorContext):
    """Stage 1: Resolve assets and load configuration."""
    # Check if direct data was provided via generation options
    if ctx.options.explicit_config_data is not None:
        ctx.config_loader = ConfigStore(
            ctx.options.explicit_config_data,
            ctx.options.explicit_template_json_data
        )
        if ctx.options.explicit_template_xlsx_bytes:
            ctx.template_xlsx_bytes = ctx.options.explicit_template_xlsx_bytes
        return

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

    # Validation: need either paths or direct data
    has_direct_data = assets and assets.config_data is not None
    if not has_direct_data and ('config' not in ctx.paths or 'template' not in ctx.paths):
         raise FileNotFoundError(f"Could not resolve config/template for '{ctx.input_path.name}'")

    # B. Load Config
    try:
        if has_direct_data:
            ctx.config_loader = ConfigStore(assets.config_data, assets.template_json_data)
        else:
            config_data, template_data = ConfigFileReader.load(ctx.paths['config'])
            ctx.config_loader = ConfigStore(config_data, template_data)
    except Exception as e:
        raise RuntimeError(f"Failed to load configuration: {e}") from e

    # Store xlsx bytes for downstream use (unknown sheet injection)
    if has_direct_data and assets.template_xlsx_bytes:
        ctx.template_xlsx_bytes = assets.template_xlsx_bytes


def _prepare_workbooks(ctx: GeneratorContext):
    """Stage 2: Load Template and Build Output Workbook."""
    ctx.output_path.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Initializing clean output workbook (JSON-only mode)")
    
    # REQUIRED: Load JSON template config
    json_config = ctx.config_loader.get_template_json_config()
    if not json_config:
        error_msg = f"CRITICAL: No JSON template found for client/config. JSON templates are now REQUIRED. (Missing *_template.json?)"
        logger.critical(error_msg)
        raise ValueError(error_msg)

    # Initialize output workbook directly from template if available (preserves native static sheets & print areas)
    from io import BytesIO
    template_xlsx_bytes = getattr(ctx, 'template_xlsx_bytes', None)
    template_path = ctx.paths.get('template')
    ctx.output_workbook = None

    if template_xlsx_bytes:
        try:
            logger.info("Loading output workbook directly from template xlsx bytes")
            ctx.output_workbook = openpyxl.load_workbook(BytesIO(template_xlsx_bytes))
        except Exception as e:
            logger.warning(f"Failed to load output workbook from template bytes: {e}")
            ctx.output_workbook = None
    elif template_path and Path(template_path).exists():
        try:
            logger.info(f"Loading output workbook directly from template path: {template_path}")
            ctx.output_workbook = openpyxl.load_workbook(template_path)
        except Exception as e:
            logger.warning(f"Failed to load output workbook from template path: {e}")
            ctx.output_workbook = None

    if ctx.output_workbook is None:
        logger.info("Initializing clean output workbook (scratch fallback)")
        ctx.output_workbook = openpyxl.Workbook()
        default_ws = ctx.output_workbook.active
        if default_ws:
            ctx.output_workbook.remove(default_ws)

    # Ensure sheets defined in JSON config exist in workbook
    for sheet_name in json_config.keys():
        if sheet_name not in ctx.output_workbook.sheetnames:
            ctx.output_workbook.create_sheet(sheet_name)
            logger.info(f"Created sheet '{sheet_name}' from JSON template")

    # Reorder sheets so dynamic JSON sheets come first, followed by static template sheets
    sheet_map = {ws.title: ws for ws in ctx.output_workbook.worksheets}
    reordered_sheets = []
    for sheet_name in json_config.keys():
        if sheet_name in sheet_map:
            reordered_sheets.append(sheet_map[sheet_name])
    for ws in ctx.output_workbook.worksheets:
        if ws.title not in json_config:
            reordered_sheets.append(ws)

    ctx.output_workbook._sheets = reordered_sheets
    logger.info(f"Reordered sheet tabs: {[ws.title for ws in ctx.output_workbook.worksheets]}")
        
    # WARNING: template_workbook is aliased to output_workbook (same object).
    # No processor currently reads from template_workbook, so this is safe.
    # Do NOT read from template_workbook expecting pristine/unmutated template data.
    ctx.template_workbook = ctx.output_workbook



    # Deep Sheet Injection
    try:
        DeepSheetBuilder.build(ctx.output_workbook, ctx.invoice_data)
    except Exception as e:
        logger.error(f"DeepSheet injection failed: {e}", exc_info=True)


def _process_sheets(ctx: GeneratorContext, session: GenerationSession):
    """Stage 3: Iterate and process each configured sheet."""
    sheets_config = ctx.config_loader.get_sheets_to_process()
    sheets_to_process = [s for s in sheets_config if s in ctx.output_workbook.sheetnames]
    
    if not sheets_to_process:
        raise ValueError("No valid sheets found to process.")

    proc_args = ctx.options

    for sheet_name in sheets_to_process:
        logger.info(f"Processing sheet '{sheet_name}'")
        try:
            tmpl_ws = ctx.template_workbook[sheet_name]
            out_ws = ctx.output_workbook[sheet_name]
            sheet_conf = ctx.config_loader.get_sheet_config(sheet_name)
            
            # Resolve active mode for config-driven source routing
            mode = "standard"
            if proc_args:
                if getattr(proc_args, 'DAF', False): mode = "daf"
                elif getattr(proc_args, 'custom', False): mode = "custom"
            ds_type = ctx.config_loader.get_data_source_type(sheet_name, mode=mode)
            
            if not ds_type:
                continue

            io_ctx = ExcelIOContext(
                template_workbook=ctx.template_workbook,
                output_workbook=ctx.output_workbook,
                template_worksheet=tmpl_ws,
                output_worksheet=out_ws
            )
            config_ctx = SheetConfigContext(
                sheet_name=sheet_name,
                sheet_config=sheet_conf,
                data_source_indicator=ds_type,
                config_loader=ctx.config_loader
            )
            data_ctx = RuntimeDataContext(
                invoice_data=ctx.invoice_data,
                cli_args=proc_args
            )
            proc_ctx = ProcessorContext(io=io_ctx, config=config_ctx, data=data_ctx)

            processor = _get_processor(ds_type, proc_ctx)

            if processor and processor.process():
                 session.log_success(sheet_name)
                 if hasattr(processor, 'header_info'):
                     session.update_logs(header_info=processor.header_info)
            else:
                 session.log_failure(sheet_name, error=RuntimeError("Processor returned False"))

        except Exception as e:
            session.log_failure(sheet_name, error=e)
            raise e


def _get_processor(ds_type: str, ctx: ProcessorContext):
    """Factory method for processors."""
    if ds_type in ["processed_tables_multi", "processed_tables", "detail_packing_list"]:
        return MultiTableProcessor(ctx)
    elif "aggregation" in ds_type or ds_type in ["DAF_aggregation", "summary_packing_list"]:
        # Fallback to aggregation for unknown/custom types that have 'aggregation' in the name
        return SingleTableProcessor(ctx)
    else:
        logger.warning(f"Unknown data source type '{ds_type}', falling back to SingleTableProcessor")
        return SingleTableProcessor(ctx)


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
    parser.add_argument("--no-auto-fit", action="store_true", help="Disable auto-fit column dimensions")
    parser.add_argument("--split-sheets", action="store_true", help="Split output workbook into individual files per sheet")
    parser.add_argument("--debug", action="store_true", help="Debug logging")
    
    args = parser.parse_args()
    
    # Configure Logging for CLI using centralized logger
    from core.logger_config import setup_logging
    from core.system_config import sys_config
    level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(log_dir=sys_config.run_log_dir, level=level)
    
    # Determine Output Path
    if args.output:
        output_path = Path(args.output)
    else:
        # Derive from input stem, output to current working directory
        input_stem = Path(args.input_data_file).stem
        output_path = Path.cwd() / f"{input_stem}.xlsx"
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

        cli_options = GenerationOptions(
            daf_mode=args.DAF,
            custom_mode=args.custom,
            enable_auto_fit=not args.no_auto_fit,
            split_sheets=args.split_sheets
        )

        paths = InvoicePathConfig(
            input_data_path=Path(args.input_data_file),
            output_path=output_path,
            template_dir=Path(args.templatedir) if args.templatedir else None,
            config_dir=Path(args.configdir) if args.configdir else None
        )
        overrides = ExplicitOverrides(
            explicit_config_path=Path(args.config) if args.config else None,
            explicit_template_path=Path(args.template) if args.template else None,
            input_data_dict=cli_data
        )
        req = InvoiceGenerationRequest(paths=paths, overrides=overrides, options=cli_options)

        run_invoice_generation(req)
        print(f"Successfully generated: {output_path}")
    except Exception as e:
        print(f"Generation failed: {e}")
        # The monitor would have already written the metadata file with the stack trace
        sys.exit(1)

if __name__ == "__main__":
    main()