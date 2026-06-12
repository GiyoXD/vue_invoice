"""
Auto Bundle Generator - Main orchestrator for automatic config generation.

This is the main entry point that:
1. Takes an Excel template file OR old config file
2. Analyzes its structure
3. Generates a complete bundle config
4. Saves it to the config_bundled directory or returns it in memory
"""

import logging
import json
import shutil
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
from io import BytesIO
from dataclasses import dataclass, field

import openpyxl

from core.utils.clock import now as ict_now, timestamp as ict_timestamp
from core.utils.pipeline_monitor import PipelineMonitor
from core.utils.snitch import snitch
from core.utils.loop_profiler import loop_profiler
from core.system_config import sys_config

from .internal.scanner import WorkbookManager, TemplateAnalysisResult
from .internal.builder import ConfigBuilder
from .internal.validator import ConfigValidator
from .internal.sanitizer import ExcelTemplateSanitizer
from .schema import BlueprintSchema

logger = logging.getLogger(__name__)


@dataclass
class BlueprintGenerationOptions:
    output_dir: Optional[str] = None
    dry_run: bool = False
    custom_prefix: Optional[str] = None
    runtime_mappings: Optional[Dict[str, str]] = None
    bundle_dir_name: Optional[str] = None
    pricing_mode: str = "standard"
    ignore_missing_description: bool = False
    in_memory: bool = False
    existing_template_json: Optional[Dict[str, Any]] = None


class BlueprintGenerator:
    """
    Main class for automatic blueprint (config + template) generation.

    Usage:
        generator = BlueprintGenerator()
        config_path = generator.generate("path/to/template.xlsx")
    """

    def __init__(self, output_base_dir: Optional[Path] = None):
        """
        Initialize the generator.

        Args:
            output_base_dir: Base directory for config output.
                           Defaults to invoice_generator/src/config_bundled/
        """
        self.scanner = WorkbookManager()
        self.builder = ConfigBuilder()
        self.validator = ConfigValidator()

        # Set output directory
        if output_base_dir:
            self.output_base_dir = Path(output_base_dir)
        else:
            self.output_base_dir = sys_config.registry_dir
            
        # Set Logger
        self.logger = logging.getLogger(self.__class__.__name__)
        
    def _load_mapping_config(self) -> Dict[str, Any]:
        """Load the user-defined mapping configuration."""
        try:
            from core.database.db_manager import get_global_mapping_config
            return get_global_mapping_config()
        except Exception as e:
            self.logger.warning(f"Failed to load mapping config from DB: {e}")
        return {}

    def analyze(self, template_path: str, legacy_format: bool = True, ignore_missing_description: bool = False) -> str:
        """
        Analyze template and return JSON string (for frontend integration).
        """
        path = Path(template_path)
        if not path.exists():
             raise FileNotFoundError(f"Template not found: {template_path}")
             
        mapping_config = self._prepare_mapping_config(None, ignore_missing_description)
        BlueprintSchema.load_dynamic_columns(mapping_config)
        
        analysis = self.scanner.scan_template(str(path), mapping_config=mapping_config)
        result = json.dumps(analysis.to_legacy_dict(), indent=2, ensure_ascii=False)
        
        # --- Profiler Report ---
        loop_profiler.report(title=f"Blueprint Analyze Profiler — {path.name}")
        loop_profiler.reset()
        
        return result
    
    def generate(self, template_path: str, options: Optional[BlueprintGenerationOptions] = None) -> Optional[Any]:
        """
        Generate bundle config from template.
        """
        if options is None:
            options = BlueprintGenerationOptions()

        path = Path(template_path)
        if not path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        self.logger.info(f"Starting blueprint generation flow for: {path.name}")
        
        # Step 0: Load workbook
        wb = self._load_workbook_file(path)
        
        # Step 1: Prepare mapping config and scan
        mapping_config = self._prepare_mapping_config(options.runtime_mappings, options.ignore_missing_description)
        BlueprintSchema.load_dynamic_columns(mapping_config)
        
        self.logger.info("Scanning template structure...")
        analysis = self.scanner.scan_template(str(path), mapping_config=mapping_config, workbook=wb)
        
        # Log scan warnings
        if analysis.warnings:
            self.logger.warning("Blueprint Generation Warnings:")
            for msg in analysis.warnings:
                self.logger.warning(f"  - {msg}")

        self._print_analysis_summary(analysis)
        
        # Step 1b: Validate column mapping conflicts
        self.logger.info("Validating column mappings for conflicts...")
        self._validate_column_mappings(analysis)
        
        # Step 2: Build bundle config
        self.logger.info("Building blueprint configuration...")
        bundle = self._assemble_blueprint_bundle(analysis, options.pricing_mode, options.custom_prefix)
        
        # Step 2b: Validate Config Structure
        self.logger.info("Validating configuration structure...")
        self._validate_bundle_structure(bundle)
        
        # Step 3: Compile layout and clean template assets
        self.logger.info("Sanitizing Excel template...")
        template_json_data, template_xlsx_bytes = self._compile_template_assets(
            wb, analysis, path, options.existing_template_json
        )
        
        # Step 4: Handle Output Mode
        if options.in_memory:
            loop_profiler.report(title=f"Blueprint Generate Profiler (In-Memory) — {path.name}")
            loop_profiler.reset()
            return bundle, template_json_data, template_xlsx_bytes

        if options.dry_run:
            self.logger.info("\n[Dry Run] Generated config:")
            print(json.dumps(bundle, indent=2, ensure_ascii=False))
            return None
            
        effective_prefix = options.custom_prefix if options.custom_prefix else analysis.customer_code
        config_file_path = self._write_blueprint_to_disk(
            bundle, template_json_data, template_xlsx_bytes, 
            effective_prefix, options.bundle_dir_name, options.output_dir
        )
        
        # --- Profiler Report ---
        loop_profiler.report(title=f"Blueprint Generate Profiler — {path.name}")
        loop_profiler.reset()
        
        return config_file_path

    def _load_workbook_file(self, template_path: Path) -> openpyxl.Workbook:
        """Load Workbook from Excel template path."""
        self.logger.info(f"Loading template workbook: {template_path.name}")
        try:
            return openpyxl.load_workbook(template_path, data_only=False)
        except Exception as e:
            self.logger.error(f"Failed to load workbook: {e}")
            raise e

    def _prepare_mapping_config(self, runtime_mappings: Optional[Dict[str, str]], ignore_missing_description: bool) -> Dict[str, Any]:
        """Load and merge global mapping configurations with user runtime mappings."""
        try:
            mapping_config = self._load_mapping_config()
        except Exception as e:
            self.logger.error(f"Failed to load mapping config: {e}")
            raise e
            
        if ignore_missing_description:
            mapping_config["ignore_missing_description"] = True
            
        if runtime_mappings:
            self.logger.info(f"   Using {len(runtime_mappings)} runtime column mappings: {runtime_mappings}")
            if "header_text_mappings" not in mapping_config:
                mapping_config["header_text_mappings"] = {"mappings": {}}
            if "mappings" not in mapping_config["header_text_mappings"]:
                 mapping_config["header_text_mappings"]["mappings"] = {}
            mapping_config["header_text_mappings"]["mappings"].update(runtime_mappings)
            
        return mapping_config

    def _validate_column_mappings(self, analysis: TemplateAnalysisResult):
        """Validate column mappings for duplicates or overlapping definitions across sheets."""
        conflict_details = []
        for sheet in analysis.sheets:
            id_to_headers = {}
            for col in sheet.columns:
                if not col.children:
                    if col.id not in id_to_headers:
                        id_to_headers[col.id] = set()
                    id_to_headers[col.id].add(col.header.strip().lower())
                
                for child in col.children:
                    if child.id not in id_to_headers:
                        id_to_headers[child.id] = set()
                    id_to_headers[child.id].add(child.header.strip().lower())
                    
            for col_id, headers in id_to_headers.items():
                if len(headers) > 1 and not col_id.startswith("col_unknown") and col_id != "col_static":
                    msg = f"Sheet '{sheet.name}': ID '{col_id}' is mapped to distinct headers {headers}"
                    self.logger.error(f"❌ MAPPING CONFLICT: {msg}")
                    conflict_details.append(msg)
                    
        if conflict_details:
            details_str = " | ".join(conflict_details)
            raise ValueError(f"Mapping conflicts detected: {details_str}. Please fix global mappings via the database or the template.")

    def _assemble_blueprint_bundle(self, analysis: TemplateAnalysisResult, pricing_mode: str, custom_prefix: Optional[str]) -> Dict[str, Any]:
        """Compile scanned template analysis into standard blueprint configuration dict."""
        bundle = self.builder.build_config(analysis)
        bundle["table_info"] = self._build_table_info(analysis)
        
        effective_prefix = custom_prefix if custom_prefix else analysis.customer_code
        if "_meta" in bundle:
            bundle["_meta"]["customer"] = effective_prefix
            bundle["_meta"]["pricing_mode"] = pricing_mode
            
        return bundle

    def _validate_bundle_structure(self, bundle: dict):
        """Validate config structure against default master specifications."""
        validation_errors = self.validator.validate(bundle)
        if validation_errors:
            self.logger.warning("⚠️  Config Validation Warnings (Deviation from Ideal Master Config):")
            for err in validation_errors:
                self.logger.warning(f"   ❌ ISSUE:  {err.get('issue')}")
                self.logger.warning(f"      DETAIL: {err.get('detail')}")
                self.logger.warning(f"      FIX:    {err.get('fix')}")
        else:
             self.logger.info("✅ Config validation passed (Matches Ideal Structure).")

    def _compile_template_assets(self, wb: openpyxl.Workbook, analysis: TemplateAnalysisResult, 
                                 template_path: Path, existing_template_json: Optional[Dict[str, Any]] = None) -> Tuple[Dict[str, Any], bytes]:
        """Sanitize workbook template and build template JSON layout metadata and XLSX bytes."""
        layout_metadata = {}
        for sheet_analysis in analysis.sheets:
            if sheet_analysis.static_layout:
                layout_metadata[sheet_analysis.name] = sheet_analysis.static_layout
            else:
                self.logger.warning(f"  Missing static layout for sheet: {sheet_analysis.name}")

        sanitizer = ExcelTemplateSanitizer()
        cleaned_wb = sanitizer.sanitize_template(wb, [sheet.name for sheet in analysis.sheets])
        
        preserved_notes = self._preserve_user_overrides(None, layout_metadata, old_data=existing_template_json)
        
        fingerprint = {
            "source_file": template_path.name,
            "created_at": ict_timestamp()
        }
        
        template_json_data = {
            "fingerprint": fingerprint,
            "template_layout": layout_metadata
        }
        if preserved_notes:
            template_json_data["notes"] = preserved_notes
            
        virtual_file = BytesIO()
        try:
            cleaned_wb.save(virtual_file)
            template_xlsx_bytes = virtual_file.getvalue()
        except Exception as e:
            self.logger.error(f"Failed to save cleaned template to memory: {e}")
            with open(template_path, "rb") as f:
                template_xlsx_bytes = f.read()
                
        return template_json_data, template_xlsx_bytes

    def _write_blueprint_to_disk(self, bundle: dict, template_json_data: dict, template_xlsx_bytes: bytes,
                                 effective_prefix: str, bundle_dir_name: Optional[str], output_dir: Optional[str]) -> Path:
        """Write compiled blueprint bundle, layout metadata, and XLSX template files to filesystem."""
        output_base = Path(output_dir) if output_dir else self.output_base_dir
        dir_name = bundle_dir_name if bundle_dir_name else effective_prefix
        config_dir = output_base / dir_name
        config_dir.mkdir(parents=True, exist_ok=True)
        
        variant_suffixes = ("_KH", "_VN")
        file_prefix = effective_prefix if effective_prefix.upper().endswith(variant_suffixes) else f"{effective_prefix}_KH"
        
        config_file = config_dir / f"{file_prefix}_config.json"
        template_config_file = config_dir / f"{file_prefix}_template.json"
        template_xlsx = config_dir / f"{file_prefix}.xlsx"
        
        # Merge overrides if file already exists
        preserved_notes = self._preserve_user_overrides(template_config_file, template_json_data["template_layout"])
        if preserved_notes:
            template_json_data["notes"] = preserved_notes
            
        # Save Template Config JSON
        self.logger.info(f"Saving template configuration to: {template_config_file.name}")
        with open(template_config_file, 'w', encoding='utf-8') as f:
            json.dump(template_json_data, f, indent=2, ensure_ascii=False)
            
        # Save Template XLSX
        try:
            with open(template_xlsx, 'wb') as f:
                f.write(template_xlsx_bytes)
            self.logger.info(f"Saved cleaned template workbook to: {template_xlsx.name}")
        except Exception as e:
            self.logger.error(f"Failed to save template file: {e}")
            
        # Save Config JSON
        self.logger.info(f"Saving blueprint configuration to: {config_file.name}")
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(bundle, f, indent=2, ensure_ascii=False)
            
        self.logger.info(f"Blueprint files saved successfully in directory: {config_dir}")
        
        return config_file

    def _preserve_user_overrides(self, template_config_file: Optional[Path], layout_metadata: Dict[str, Any], old_data: Optional[Dict[str, Any]] = None) -> Optional[Any]:
        """
        Preserve user overrides from an existing template config.
        
        If an old template JSON exists, carry over any mode-dependent
        overrides (dict values in template_header_content / template_footer_rows,
        e.g. {"default":"X","standard":"Y","daf":"Z"}).
        These are user-entered via Template Inspector and would be lost on regeneration.
        Also preserves the 'notes' field.
        """
        if old_data is None:
            if template_config_file is None or not template_config_file.exists():
                return None
            try:
                with open(template_config_file, 'r', encoding='utf-8') as f:
                    old_data = json.load(f)
            except Exception as e:
                self.logger.warning(f"   [Override Preservation] Could not read old config file: {e}")
                return None
        
        try:
            preserved_notes = old_data.get("notes")
            old_layout = old_data.get("template_layout", {})
            
            override_count = 0
            for sheet_name, old_sheet in old_layout.items():
                if sheet_name not in layout_metadata:
                    continue
                
                # --- HEADER CONTENT OVERRIDES ---
                old_hc = old_sheet.get("template_header_content") or old_sheet.get("header_content", {})
                new_hc = layout_metadata[sheet_name].get("template_header_content") or layout_metadata[sheet_name].get("header_content", {})
                
                for cell_addr, old_val in old_hc.items():
                    if isinstance(old_val, dict):
                        new_plain = new_hc.get(cell_addr)
                        preserved_override = {
                            "default": new_plain if (new_plain is not None and not isinstance(new_plain, dict)) else ""
                        }
                        if "standard" in old_val:
                            preserved_override["standard"] = old_val["standard"]
                        if "daf" in old_val:
                            preserved_override["daf"] = old_val["daf"]
                        new_hc[cell_addr] = preserved_override
                        override_count += 1
                
                if "template_header_content" in layout_metadata[sheet_name]:
                    layout_metadata[sheet_name]["template_header_content"] = new_hc
                else:
                    layout_metadata[sheet_name]["header_content"] = new_hc
                
                # --- FOOTER ROW OVERRIDES ---
                old_footer_rows = old_sheet.get("template_footer_rows") or old_sheet.get("footer_rows", [])
                new_footer_rows = layout_metadata[sheet_name].get("template_footer_rows") or layout_metadata[sheet_name].get("footer_rows", [])
                
                old_footer_overrides = {}
                for old_row in old_footer_rows:
                    rel_idx = old_row.get("relative_index")
                    for old_cell in old_row.get("cells", []):
                        old_cell_val = old_cell.get("value")
                        if isinstance(old_cell_val, dict):
                            col_idx = old_cell.get("col_index")
                            old_footer_overrides[(rel_idx, col_idx)] = old_cell_val
                
                if old_footer_overrides:
                    for new_row in new_footer_rows:
                        rel_idx = new_row.get("relative_index")
                        for new_cell in new_row.get("cells", []):
                            col_idx = new_cell.get("col_index")
                            key = (rel_idx, col_idx)
                            if key in old_footer_overrides:
                                old_override = old_footer_overrides[key]
                                new_plain = new_cell.get("value")
                                preserved_override = {
                                    "default": new_plain if (new_plain is not None and not isinstance(new_plain, dict)) else ""
                                }
                                if "standard" in old_override:
                                    preserved_override["standard"] = old_override["standard"]
                                if "daf" in old_override:
                                    preserved_override["daf"] = old_override["daf"]
                                new_cell["value"] = preserved_override
                                override_count += 1
            
            if override_count > 0:
                self.logger.info(f"   [Override Preservation] Merged {override_count} user overrides from existing template.")
            else:
                self.logger.info("   [Override Preservation] No user overrides found to preserve.")
            
            return preserved_notes
        except Exception as e:
            self.logger.warning(f"   [Override Preservation] Could not merge old overrides: {e}")
            return None
    
    def _build_table_info(self, analysis: TemplateAnalysisResult) -> Dict[str, Any]:
        """
        Build the table_info summary index from the analysis result.
        
        This is a flat summary of key template metadata for quick API access.
        """
        fallback_description = None
        hs_code = None

        sheets_ordered = sorted(
            analysis.sheets,
            key=lambda s: s.name.lower() == "packing list",
            reverse=True
        )
        for sheet in sheets_ordered:
            if not fallback_description and sheet.static_content_hints:
                fallback_description = sheet.static_content_hints.get("description_fallback")
            if not hs_code and sheet.footer_info and sheet.footer_info.hs_code_text:
                hs_code = sheet.footer_info.hs_code_text
            if fallback_description and hs_code:
                break

        return {
            "fallback_description": {"standard": fallback_description, "daf": fallback_description} if fallback_description else None,
            "hs_code": hs_code
        }

    def _print_analysis_summary(self, analysis: TemplateAnalysisResult):
        """Print summary of template analysis."""
        self.logger.info(f"\n   Customer Code: {analysis.customer_code}")
        self.logger.info(f"   Sheets found: {len(analysis.sheets)}")
        
        for sheet in analysis.sheets:
            self.logger.info(f"\n   [{sheet.name}]")
            self.logger.info(f"      Header row: {sheet.header_row}")
            self.logger.info(f"      Data source: {sheet.data_source}")
            self.logger.info(f"      Columns: {len(sheet.columns)}")
            
            for col in sheet.columns:
                children_info = f" ({len(col.children)} children)" if col.children else ""
                self.logger.info(f"         - {col.id}: '{col.header}'{children_info}")
