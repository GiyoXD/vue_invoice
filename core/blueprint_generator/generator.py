"""
Auto Bundle Generator - Main orchestrator for automatic config generation.

This is the main entry point that:
1. Takes an Excel template file OR old config file
2. Analyzes its structure
3. Generates a complete bundle config
4. Saves it to the config_bundled directory
"""

import logging
import json
import argparse
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
from core.utils.clock import now as ict_now, timestamp as ict_timestamp
import openpyxl


from .internal.scanner import ExcelLayoutScanner, TemplateAnalysisResult
from .internal.builder import ConfigBuilder
from .internal.validator import ConfigValidator
from core.utils.pipeline_monitor import PipelineMonitor
from core.utils.snitch import snitch
from core.utils.loop_profiler import loop_profiler


logger = logging.getLogger(__name__)


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
        self.scanner = ExcelLayoutScanner()
        self.builder = ConfigBuilder()
        self.validator = ConfigValidator()

        from core.system_config import sys_config        
        # Set output directory
        if output_base_dir:
            self.output_base_dir = Path(output_base_dir)
        else:
            self.output_base_dir = sys_config.registry_dir
            
        # Set Mapping config path
        self.mapping_config_path = sys_config.mapping_config_path
        
        self.logger = logging.getLogger(self.__class__.__name__)
        
    def _load_mapping_config(self) -> Dict[str, Any]:
        """Load the user-defined mapping configuration."""
        if self.mapping_config_path.exists():
            try:
                with open(self.mapping_config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                self.logger.warning(f"Failed to load mapping config: {e}")
        return {}

    def analyze(self, template_path: str, legacy_format: bool = True) -> str:
        """
        Analyze template and return JSON string (for frontend integration).
        """
        template_path = Path(template_path)
        if not template_path.exists():
             raise FileNotFoundError(f"Template not found: {template_path}")
             
        mapping_config = self._load_mapping_config()
        analysis = self.scanner.scan_template(str(template_path), mapping_config=mapping_config)
        
        # Currently, legacy_format is always assumed or the output format is identical
        result = json.dumps(analysis.to_legacy_dict(), indent=2, ensure_ascii=False)
        
        # --- Profiler Report ---
        loop_profiler.report(title=f"Blueprint Analyze Profiler — {template_path.name}")
        loop_profiler.reset()
        
        return result
    
    def generate(self, template_path: str, output_dir: Optional[str] = None,
                 dry_run: bool = False, monitor: Optional[PipelineMonitor] = None,
                 custom_prefix: Optional[str] = None,
                 runtime_mappings: Optional[Dict[str, str]] = None,
                 bundle_dir_name: Optional[str] = None,
                 pricing_mode: str = "standard") -> Optional[Path]:
        """
        Generate bundle config from template.
        
        Args:
            template_path: Path to Excel template file
            output_dir: Optional custom output directory
            dry_run: If True, print config but don't save
            monitor: Optional pipeline monitor for logging
            custom_prefix: Optional custom prefix to use instead of auto-detected customer code
            runtime_mappings: Optional dict of {header_text: col_id} to override/add to global mappings
            bundle_dir_name: Optional folder name override. When set, output folder uses this name
                           instead of the prefix (e.g. folder='MOTO' but files='MOTO_KH_config.json')
            pricing_mode: Pricing calculation mode. 'standard' (sqft × price) or 'net' (net_weight × global price)
            
        Returns:
            Path to generated config file, or None if dry_run
        """
        template_path = Path(template_path)
        
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        self.logger.info("=" * 60)
        self.logger.info("Auto Bundle Generator")
        self.logger.info("=" * 60)
        self.logger.info("Template: {template_path}")
        
        
        # Step 0: Load Workbook ONCE (Optimization)
        self.logger.info("\n[Step 0] Loading workbook...")
        try:
            wb = openpyxl.load_workbook(template_path, data_only=False)
        except Exception as e:
            self.logger.error(f"Failed to load workbook: {e}")
            raise e
        
        # Step 1: Analyze template
        self.logger.info("\n[Step 1] Scanning template structure...")
        try:
            mapping_config = self._load_mapping_config()
        except Exception as e:
            self.logger.error(f"Failed to load mapping config: {e}")
            raise e
        
        # Inject Runtime Mappings (from API/User)
        if runtime_mappings:
            self.logger.info(f"   Using {len(runtime_mappings)} runtime column mappings: {runtime_mappings}")
            if "header_text_mappings" not in mapping_config:
                mapping_config["header_text_mappings"] = {"mappings": {}}
            if "mappings" not in mapping_config["header_text_mappings"]:
                 mapping_config["header_text_mappings"]["mappings"] = {}
            
            # Update the config used for scanning
            mapping_config["header_text_mappings"]["mappings"].update(runtime_mappings)
            
            # [Smart Feature] "One-Shot Learning": Save new mappings globally
            try:
                from core.system_config import sys_config
                config_path = sys_config.mapping_config_path
                
                # Ensure directory exists
                config_path.parent.mkdir(parents=True, exist_ok=True)
                
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(mapping_config, f, indent=4)
                self.logger.info(f"   [Learning] Saved {len(runtime_mappings)} new mappings to global config at {config_path}.")
            except Exception as e:
                self.logger.warning(f"   [Learning Failed] Could not save mappings to disk: {e}")

        analysis = self.scanner.scan_template(str(template_path), mapping_config=mapping_config, workbook=wb)
        
        # [Proactive Warning Reporting]
        if analysis.warnings:
            self.logger.warning("\n" + "!" * 60)
            self.logger.warning("!!! BLUEPRINT GENERATION WARNINGS !!!")
            for msg in analysis.warnings:
                self.logger.warning(f"  {msg}")
            self.logger.warning("!" * 60 + "\n")

        self._print_analysis_summary(analysis)
        
        # Use custom prefix if provided, otherwise use detected customer code
        effective_prefix = custom_prefix if custom_prefix else analysis.customer_code
        if custom_prefix:
            self.logger.info(f"\n   Using custom prefix: {effective_prefix}")
            
        # Step 1b: Validate mappings for conflicts
        self.logger.info("\n[Step 1b] Validating mappings for conflicts...")
        conflict_details = []
        for sheet in analysis.sheets:
            id_to_headers = {}
            for col in sheet.columns:
                # If a column has children, we shouldn't consider its own header mapping as a conflict 
                # against its children, because it's just a grouping label (e.g. "Quantity" over "PCS" and "SF")
                if not col.children:
                    if col.id not in id_to_headers:
                        id_to_headers[col.id] = set()
                    # Store the stripped lowercase version to check for ACTUAL text differences
                    id_to_headers[col.id].add(col.header.strip().lower())
                
                # Do the same for children
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
            raise ValueError(f"Mapping conflicts detected: {details_str}. Please fix mapping_config.json or the template.")

        # Step 2: Build bundle config
        self.logger.info("\n[Step 2] Building blueprint config...")
        bundle = self.builder.build_config(analysis)
        
        # Inject table_info as a top-level summary index for quick API access.
        # The invoice renderer reads from the deep paths (data_flow, footer.add_ons);
        # the API reads from this single flat key — no path-digging needed.
        bundle["table_info"] = self._build_table_info(analysis)
        
        # Update bundle metadata with effective prefix
        if custom_prefix and "_meta" in bundle:
            bundle["_meta"]["customer"] = custom_prefix
        
        # Write pricing_mode to _meta
        if "_meta" in bundle:
            bundle["_meta"]["pricing_mode"] = pricing_mode

        # Step 2b: Validate Config Structure
        self.logger.info("\n[Step 2b] Validating config structure...")
        validation_errors = self.validator.validate(bundle)
        if validation_errors:
            self.logger.warning("⚠️  Config Validation Warnings (Deviation from Ideal Master Config):")
            for err in validation_errors:
                self.logger.warning(f"   ------------------------------------------------------------")
                self.logger.warning(f"   ❌ ISSUE:  {err.get('issue')}")
                self.logger.warning(f"      DETAIL: {err.get('detail')}")
                self.logger.warning(f"      FIX:    {err.get('fix')}")
            self.logger.warning(f"   ------------------------------------------------------------\n")
        else:
             self.logger.info("✅ Config validation passed (Matches Ideal Structure).")
        
        
        # Step 3: Save or print
        if dry_run:
            self.logger.info("\n[Dry Run] Generated config:")
            print(json.dumps(bundle, indent=2, ensure_ascii=False))
            return None
        
        # Determine output path using effective prefix
        if output_dir:
            output_base = Path(output_dir)
        else:
            output_base = self.output_base_dir
        
        dir_name = bundle_dir_name if bundle_dir_name else effective_prefix
        config_dir = output_base / dir_name
        
        # Default to _KH suffix unless prefix already ends with a variant suffix
        variant_suffixes = ("_KH", "_VN")
        file_prefix = effective_prefix if effective_prefix.upper().endswith(variant_suffixes) else f"{effective_prefix}_KH"
        config_file = config_dir / f"{file_prefix}_config.json"
        
        # Create directory
        config_dir.mkdir(parents=True, exist_ok=True)

        # Step 3: Generate Clean Template
        if not dry_run:
            template_path_res, layout_metadata = self._generate_clean_template(
                template_path, analysis, config_dir, monitor, effective_prefix, workbook=wb
            )
            
            # SAVE SEPARATE TEMPLATE CONFIG
            template_config_file = config_dir / f"{file_prefix}_template.json"
            
            # [Preserve User Overrides]
            preserved_notes = self._preserve_user_overrides(template_config_file, layout_metadata)
            
            # [Fingerprint]
            fingerprint = {
                "source_file": template_path.name,
                "created_at": ict_timestamp()
            }
            

            
            self.logger.info(f"   Saving Template Config: {template_config_file.name}")
            template_json_data = {
                "fingerprint": fingerprint,
                "template_layout": layout_metadata
            }
            if preserved_notes:
                template_json_data["notes"] = preserved_notes
                
            with open(template_config_file, 'w', encoding='utf-8') as f:
                json.dump(template_json_data, f, indent=2, ensure_ascii=False)
                
            # Note: We do NOT inject it into the main bundle config anymore.
        
        self.logger.info(f"\n[Step 4] Saving config to: {config_file}")
        
        # Write config
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(bundle, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"✅ Config saved successfully!")
        self.logger.info(f"   Directory: {config_dir}")
        self.logger.info(f"   File: {config_file.name}")
        
        # --- Profiler Report ---
        loop_profiler.report(title=f"Blueprint Generate Profiler — {template_path.name}")
        loop_profiler.reset()
        
        return config_file



    def _generate_clean_template(self, template_path: Path, analysis: TemplateAnalysisResult, 
                                 output_dir: Path, monitor: Optional[PipelineMonitor] = None,
                                 custom_prefix: Optional[str] = None,
                                 workbook: Optional[openpyxl.Workbook] = None) -> Tuple[Path, Dict[str, Any]]:
        """
        Generate a clean template from the raw input file.
        
        Sanitizes the template (strips data rows) but does NOT process images
        since that was causing file corruption.
        
        Args:
            template_path: Path to raw Excel file
            analysis: Analysis result
            output_dir: Directory to save template
            monitor: Optional pipeline monitor
            custom_prefix: Optional custom prefix for naming
            workbook: Optional pre-loaded workbook (optimization)
            
        Returns:
            Tuple of (Path to saved template, layout_metadata)
        """
        import openpyxl
        from .internal.sanitizer import ExcelTemplateSanitizer
        
        self.logger.info(f"\n[Step 3] Sanitizing Template...")
        
        # Use custom prefix if provided, otherwise use detected customer code
        effective_prefix = custom_prefix if custom_prefix else analysis.customer_code
        
        # Target template file path
        # Use file_prefix (with variant suffix) if available, otherwise default to _KH
        variant_suffixes = ("_KH", "_VN")
        file_prefix = effective_prefix if effective_prefix.upper().endswith(variant_suffixes) else f"{effective_prefix}_KH"
        template_file = output_dir / f"{file_prefix}.xlsx"
        
        # Load and sanitize the template
        sanitizer = ExcelTemplateSanitizer()
        
        if workbook:
             # Use shared workbook object (Optimization)
             # NOTE: We must be careful if sanitizer modifies it in place.
             # Sanitizer DOES delete rows. If we need original state later, we might have issues.
             # But here, 'generate' is the end of the line. We don't need the original state after this.
             self.logger.info("   Using pre-loaded workbook for cleaning.")
             wb = workbook
        else:
             self.logger.info("   Loading workbook from disk...")
             wb = openpyxl.load_workbook(template_path)
        
        # Sanitize (strips data rows, no image handling)
        cleaned_wb, layout_metadata = sanitizer.sanitize_template(wb, analysis)
        
        # Save the cleaned template
        try:
            cleaned_wb.save(template_file)
            self.logger.info(f"   Cleaned Template Saved: {template_file.name}")
        except Exception as e:
            self.logger.error(f"   Failed to save cleaned template: {e}")
            if monitor:
                monitor.log_warning(f"Failed to save template: {e}")
            # Fallback: copy original file
            import shutil
            shutil.copy2(template_path, template_file)
            self.logger.info(f"   Fallback: Copied original template")
            
        return template_file, layout_metadata
    
    def _preserve_user_overrides(self, template_config_file: Path, layout_metadata: Dict[str, Any]) -> Optional[Any]:
        """
        Preserve user overrides from an existing template config.
        
        If an old template JSON exists, carry over any mode-dependent
        overrides (dict values in template_header_content / template_footer_rows,
        e.g. {"default":"X","standard":"Y","daf":"Z"}).
        These are user-entered via Template Inspector and would be lost on regeneration.
        Also preserves the 'notes' field.
        
        Args:
            template_config_file: Path to the existing template JSON file
            layout_metadata: The newly generated layout metadata (mutated in place)
            
        Returns:
            Preserved notes value, or None if no old config exists
        """
        if not template_config_file.exists():
            return None
        
        try:
            with open(template_config_file, 'r', encoding='utf-8') as f:
                old_data = json.load(f)
            
            preserved_notes = old_data.get("notes")
            old_layout = old_data.get("template_layout", {})
            
            override_count = 0
            for sheet_name, old_sheet in old_layout.items():
                if sheet_name not in layout_metadata:
                    continue
                
                # --- HEADER CONTENT OVERRIDES ---
                # Support both old key ("header_content") and new key ("template_header_content")
                old_hc = old_sheet.get("template_header_content") or old_sheet.get("header_content", {})
                new_hc = layout_metadata[sheet_name].get("template_header_content") or layout_metadata[sheet_name].get("header_content", {})
                
                for cell_addr, old_val in old_hc.items():
                    if isinstance(old_val, dict):
                        # Build fresh override: default from NEW scan, only standard/daf from OLD
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
                
                # Write back to the correct key used by the sanitizer
                if "template_header_content" in layout_metadata[sheet_name]:
                    layout_metadata[sheet_name]["template_header_content"] = new_hc
                else:
                    layout_metadata[sheet_name]["header_content"] = new_hc
                
                # --- FOOTER ROW OVERRIDES ---
                old_footer_rows = old_sheet.get("template_footer_rows") or old_sheet.get("footer_rows", [])
                new_footer_rows = layout_metadata[sheet_name].get("template_footer_rows") or layout_metadata[sheet_name].get("footer_rows", [])
                
                # Build lookup: (relative_index, col_index) -> old cell value
                old_footer_overrides = {}
                for old_row in old_footer_rows:
                    rel_idx = old_row.get("relative_index")
                    for old_cell in old_row.get("cells", []):
                        old_cell_val = old_cell.get("value")
                        if isinstance(old_cell_val, dict):
                            # This cell has mode overrides
                            col_idx = old_cell.get("col_index")
                            old_footer_overrides[(rel_idx, col_idx)] = old_cell_val
                
                # Merge old overrides into the new footer rows
                if old_footer_overrides:
                    for new_row in new_footer_rows:
                        rel_idx = new_row.get("relative_index")
                        for new_cell in new_row.get("cells", []):
                            col_idx = new_cell.get("col_index")
                            key = (rel_idx, col_idx)
                            if key in old_footer_overrides:
                                old_override = old_footer_overrides[key]
                                # Build fresh override: default from NEW scan, only standard/daf from OLD
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
        The invoice renderer reads from the deep config paths; this index is only
        for the API/UI layer.
        """
        fallback_description = None
        hs_code = None

        # Prefer 'Packing list' sheet, fall back to any sheet that has the values.
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
