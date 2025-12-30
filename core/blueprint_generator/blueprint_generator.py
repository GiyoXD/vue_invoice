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
import sys
import argparse
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
from datetime import datetime

from .excel_scanner import ExcelLayoutScanner, TemplateAnalysisResult
from .config_builder import ConfigBuilder
from .legacy_migrator import LegacyConfigMigrator
from .validator import ConfigValidator
from core.utils.pipeline_monitor import PipelineMonitor

logger = logging.getLogger(__name__)


class BlueprintGenerator:
    """
    Main class for automatic blueprint (config + template) generation.
    
    Usage:
        generator = BlueprintGenerator()
        config_path = generator.generate("path/to/template.xlsx")
        # OR convert old config
        config_path = generator.convert_old_config("path/to/old_config.json")
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
        self.migrator = LegacyConfigMigrator()
        self.validator = ConfigValidator()
        
        # Find project root for asset resolution
        current_dir = Path(__file__).parent
        # Go up to core/blueprint_generator -> core -> PROJECT_ROOT
        # Depending on where this is run, __file__ might be relative or absolute.
        # Ideally we trust the passed output_base_dir or fallback relative to this file.
        project_root = current_dir.parent.parent.parent
            
        from core.system_config import sys_config
        
        # Set output directory
        if output_base_dir:
            self.output_base_dir = Path(output_base_dir)
        else:
            self.output_base_dir = sys_config.registry_dir
            
        # Set Mapping config path (Fixed: Now properly indented)
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
        
        if legacy_format:
            return json.dumps(analysis.to_legacy_dict(), indent=2, ensure_ascii=False)
        else:
             return json.dumps(analysis.to_legacy_dict(), indent=2, ensure_ascii=False)
    
    def convert_old_config(self, config_path: str, output_dir: Optional[str] = None,
                           dry_run: bool = False) -> Optional[Path]:
        """
        Convert old-format config to bundle format.
        
        Args:
            config_path: Path to old config JSON file
            output_dir: Optional custom output directory
            dry_run: If True, print config but don't save
            
        Returns:
            Path to generated config file, or None if dry_run
        """
        config_path = Path(config_path)
        
        if not config_path.exists():
            raise FileNotFoundError(f"Config not found: {config_path}")
        
        self.logger.info(f"=" * 60)
        self.logger.info(f"Config Converter")
        self.logger.info(f"=" * 60)
        self.logger.info(f"Source: {config_path}")
        
        # Convert
        bundle = self.migrator.convert(str(config_path))
        customer_code = bundle["_meta"]["customer"]
        
        self.logger.info(f"Customer: {customer_code}")
        self.logger.info(f"Sheets: {bundle['processing']['sheets']}")
        
        if dry_run:
            self.logger.info("\n[Dry Run] Generated config:")
            print(json.dumps(bundle, indent=2, ensure_ascii=False))
            return None
        
        # Determine output path
        if output_dir:
            output_base = Path(output_dir)
        else:
            output_base = self.output_base_dir
        
        config_dir = output_base / customer_code
        config_file = config_dir / f"{customer_code}_config.json"
        
        self.logger.info(f"\nSaving to: {config_file}")
        
        config_dir.mkdir(parents=True, exist_ok=True)
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(bundle, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"✅ Config saved successfully!")
        
        return config_file
    
    def generate(self, template_path: str, output_dir: Optional[str] = None,
                 dry_run: bool = False, monitor: Optional[PipelineMonitor] = None) -> Optional[Path]:
        """
        Generate bundle config from template.
        
        Args:
            template_path: Path to Excel template file
            output_dir: Optional custom output directory
            dry_run: If True, print config but don't save
            
        Returns:
            Path to generated config file, or None if dry_run
        """
        template_path = Path(template_path)
        
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        self.logger.info(f"=" * 60)
        self.logger.info(f"Auto Bundle Generator")
        self.logger.info(f"=" * 60)
        self.logger.info(f"Template: {template_path}")
        
        
        # Step 1: Analyze template
        self.logger.info("\n[Step 1] Scanning template structure...")
        mapping_config = self._load_mapping_config()
        analysis = self.scanner.scan_template(str(template_path), mapping_config=mapping_config)
        
        self._print_analysis_summary(analysis)
        
        # Step 2: Build bundle config
        self.logger.info("\n[Step 2] Building blueprint config...")
        bundle = self.builder.build_config(analysis)

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
        
        # Determine output path
        if output_dir:
            output_base = Path(output_dir)
        else:
            output_base = self.output_base_dir
        
        config_dir = output_base / analysis.customer_code
        config_file = config_dir / f"{analysis.customer_code}_config.json"
        
        # Create directory
        config_dir.mkdir(parents=True, exist_ok=True)

        # Step 3: Generate Clean Template
        if not dry_run:
            template_path_res, layout_metadata = self._generate_clean_template(template_path, analysis, config_dir, monitor)
            
            # SAVE SEPARATE TEMPLATE CONFIG
            # User request: "create a seperate, this new config is template config"
            template_config_file = config_dir / f"{analysis.customer_code}_template.json"
            
            self.logger.info(f"   Saving Template Config: {template_config_file.name}")
            with open(template_config_file, 'w', encoding='utf-8') as f:
                json.dump({"template_layout": layout_metadata}, f, indent=2, ensure_ascii=False)
                
            # Note: We do NOT inject it into the main bundle config anymore.
        
        self.logger.info(f"\n[Step 4] Saving config to: {config_file}")
        
        # Write config
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(bundle, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"✅ Config saved successfully!")
        self.logger.info(f"   Directory: {config_dir}")
        self.logger.info(f"   File: {config_file.name}")
        
        return config_file

    def _generate_clean_template(self, template_path: Path, analysis: TemplateAnalysisResult, 
                                 output_dir: Path, monitor: Optional[PipelineMonitor] = None) -> Tuple[Path, Dict[str, Any]]:
        """
        Generate a clean template from the raw input file.
        
        Args:
            template_path: Path to raw Excel file
            analysis: Analysis result
            output_dir: Directory to save clean template
            
        Returns:
            Tuple of (Path to saved clean template, layout_metadata)
        """
        import openpyxl
        from .excel_sanitizer import ExcelTemplateSanitizer
        
        self.logger.info(f"\n[Step 3] Sanitizing Template...")
        
        sanitizer = ExcelTemplateSanitizer()
        
        # Load raw workbook
        wb = openpyxl.load_workbook(template_path)
        
        # Clean it (updated to return tuple)
        cleaned_wb, layout_metadata = sanitizer.sanitize_template(wb, analysis)
        
        # Save it
        template_file = output_dir / f"{analysis.customer_code}.xlsx"
        try:
            cleaned_wb.save(template_file)
            self.logger.info(f"   Cleaned Template: {template_file.name}")
        except Exception as e:
            msg = f"Failed to save cleaned template (likely image stream issue): {e}"
            self.logger.error(f"   {msg}")
            if monitor:
                monitor.log_warning(msg)
            self.logger.info("   Proceeding to save JSON config anyway.")
            
        return template_file, layout_metadata
    
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


def main():
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Auto-generate invoice bundle configs from Excel templates",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Generate config from CLW template
  python -m core.blueprint_generator.blueprint_generator CLW.xlsx -v

The tool will:
  1. Analyze the Excel template structure
  2. Detect sheet types (Invoice, Contract, Packing list)
  3. Extract column layouts, fonts, widths
  4. Generate a complete bundle config
  5. Save to config_bundled/{CUSTOMER}_config/
        """
    )
    
    parser.add_argument(
        'template',
        help='Path to Excel template file (e.g., CLW.xlsx)'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Output directory (default: invoice_generator/src/config_bundled/)'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Print generated config without saving'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )
    
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Minimal output'
    )
    
    parser.add_argument(
        '--convert',
        action='store_true',
        help='Convert old config format instead of analyzing template'
    )
    
    args = parser.parse_args()
    
    # Configure logging
    if args.verbose:
        log_level = logging.DEBUG
    elif args.quiet:
        log_level = logging.WARNING
    else:
        log_level = logging.INFO
    
    logging.basicConfig(
        level=log_level,
        format='%(message)s'
    )
    
    try:
        # Determine paths for Monitor
        args_template_path = Path(args.template)
        output_base_dir = Path(args.output) if args.output else None
        
        from core.system_config import sys_config
        log_dir = sys_config.run_log_dir
        log_dir.mkdir(parents=True, exist_ok=True)
        
        monitor_path = log_dir / f"{args_template_path.stem}_blueprint.json"
        
        with PipelineMonitor(monitor_path, args=args, step_name="Blueprint Generator") as monitor:
            
            generator = BlueprintGenerator(output_base_dir=output_base_dir)
            
            # Check file extension to auto-detect mode
            is_json = args_template_path.suffix.lower() == '.json'
            
            if args.convert or is_json:
                # Convert old config
                result = generator.convert_old_config(
                    args.template,
                    output_dir=args.output,
                    dry_run=args.dry_run
                )
            else:
                # Analyze template
                result = generator.generate(
                    args.template,
                    output_dir=args.output,
                    dry_run=args.dry_run,
                    monitor=monitor
                )
            
            if result:
                monitor.log_process_item("Configuration Generated", status="success")
                print(f"\n[SUCCESS] Config generated at {result}")
            
            return 0
            
    except Exception as e:
        print(f"❌ ERROR: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
