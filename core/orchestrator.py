# core/orchestrator.py
import sys
import os
from pathlib import Path
from typing import Dict, Optional, Tuple, Any

# Import the logic directly!
from core.invoice_generator.generate_invoice import run_invoice_generation
from core.data_parser.main import run_invoice_automation
from core.data_parser.data_processor import DataValidationError
from core.utils.snitch import snitch

class Orchestrator:
    """
    Service to orchestrate backend processes.
    Refactored to use direct python calls where possible.
    """

    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        
    @snitch
    def process_excel_to_json(self, excel_path, output_dir: Path, input_filename_override: str = None, ignore_tare_warning: bool = False, ignore_cbm_warning: bool = False) -> Tuple[Path, str]:
        """
        Directly calls the Data Parser library function.
        No more subprocess overhead.
        """
        try:
            # Call the refactored main function from data_parser
            # It returns (json_path, identifier) on success
            json_path, identifier = run_invoice_automation(
                input_excel_override=excel_path if hasattr(excel_path, 'read') else str(excel_path),
                input_filename_override=input_filename_override,
                output_dir_override=str(output_dir),
                ignore_tare_warning=ignore_tare_warning,
                ignore_cbm_warning=ignore_cbm_warning
            )
            return json_path, identifier

        except DataValidationError:
            # User-facing validation errors pass through cleanly (no traceback wrapping)
            raise
        except Exception as e:
            # Capture the full traceback for the UI to display
            import traceback
            tb = traceback.format_exc()
            raise RuntimeError(f"Data Parser Failed:\n{tb}") from e

    @snitch
    def generate_invoice(self, 
                        json_path: Path, 
                        output_path: Path, 
                        template_dir: Path, 
                        config_dir: Path, 
                        explicit_config_path: Path = None,
                        explicit_template_path: Path = None,
                        input_data_dict: Dict = None,
                        options = None):
        """
        Directly calls the Invoice Generator library function.
        No more subprocess overhead or serialization issues.
        """
        try:
            # CALLING DIRECTLY
            result = run_invoice_generation(
                input_data_path=json_path,
                output_path=output_path,
                template_dir=template_dir,
                config_dir=config_dir,
                explicit_config_path=explicit_config_path,
                explicit_template_path=explicit_template_path,
                input_data_dict=input_data_dict,
                options=options
            )
            return result

        except Exception as e:
            # Capture the full traceback for the UI to display
            import traceback
            tb = traceback.format_exc()
            raise RuntimeError(f"Invoice Generation Failed:\n{tb}") from e

    # --- Blueprint / Template Management ---

    def analyze_template(self, template_path: Path, legacy_format: bool = True, ignore_missing_description: bool = False) -> str:
        """
        Wraps BlueprintGenerator.analyze.
        Returns the analysis result as a JSON string.
        """
        try:
            from core.blueprint_generator import BlueprintGenerator
            
            # Re-initialize generator to ensure fresh state
            generator = BlueprintGenerator(self.project_root)
            
            return generator.analyze(str(template_path), legacy_format=legacy_format, ignore_missing_description=ignore_missing_description)
            
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            raise RuntimeError(f"Template Analysis Failed:\n{tb}") from e

    def generate_blueprint_bundle(self, 
                                template_path: Path, 
                                output_dir: Optional[Path] = None, 
                                custom_prefix: str = None,
                                runtime_mappings: Dict[str, str] = None,
                                bundle_dir_name: str = None,
                                pricing_mode: str = "standard",
                                ignore_missing_description: bool = False,
                                in_memory: bool = False,
                                existing_template_json: Optional[Dict[str, Any]] = None) -> Any:
        """
        Wraps BlueprintGenerator.generate.
        Generates the config and clean template bundle.
        """
        try:
            from core.blueprint_generator.generator import BlueprintGenerator, BlueprintGenerationOptions
            
            generator = BlueprintGenerator(self.project_root)
            
            options = BlueprintGenerationOptions(
                output_dir=str(output_dir) if output_dir else None,
                dry_run=False,
                custom_prefix=custom_prefix,
                runtime_mappings=runtime_mappings,
                bundle_dir_name=bundle_dir_name,
                pricing_mode=pricing_mode,
                ignore_missing_description=ignore_missing_description,
                in_memory=in_memory,
                existing_template_json=existing_template_json
            )
            
            result_path = generator.generate(
                template_path=str(template_path),
                options=options
            )
            
            return result_path
            
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            raise RuntimeError(f"Blueprint Generation Failed:\n{tb}") from e