# core/orchestrator.py
import sys
import os
from pathlib import Path
from typing import List, Dict, Tuple

# Import the logic directly!
from core.invoice_generator.generate_invoice import run_invoice_generation
from core.data_parser.main import run_invoice_automation

class Orchestrator:
    """
    Service to orchestrate backend processes.
    Refactored to use direct python calls where possible.
    """

    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        
    def process_excel_to_json(self, excel_path: Path, output_dir: Path) -> Tuple[Path, str]:
        """
        Directly calls the Data Parser library function.
        No more subprocess overhead.
        """
        try:
            # Call the refactored main function from data_parser
            # It returns (json_path, identifier) on success
            json_path, identifier = run_invoice_automation(
                input_excel_override=str(excel_path),
                output_dir_override=str(output_dir)
            )
            return json_path, identifier

        except Exception as e:
            # Capture the full traceback for the UI to display
            import traceback
            tb = traceback.format_exc()
            raise RuntimeError(f"Data Parser Failed:\n{tb}") from e

    def generate_invoice(self, 
                        json_path: Path, 
                        output_path: Path, 
                        template_dir: Path, 
                        config_dir: Path, 
                        flags: List[str] = None,
                        input_data_dict: Dict = None) -> Path:
        """
        Directly calls the Invoice Generator library function.
        No more subprocess overhead or serialization issues.
        """
        flags = flags or []
        
        
        # Convert legacy CLI flags to function arguments
        daf_mode = "--DAF" in flags
        custom_mode = "--custom" in flags
        
        explicit_config_path = None
        explicit_template_path = None
        
        # Simple parser for value flags
        if "--config" in flags:
            try:
                idx = flags.index("--config")
                if idx + 1 < len(flags):
                    explicit_config_path = Path(flags[idx + 1])
            except ValueError:
                pass
                
        if "--template" in flags:
            try:
                idx = flags.index("--template")
                if idx + 1 < len(flags):
                    explicit_template_path = Path(flags[idx + 1])
            except ValueError:
                pass

        try:
            # CALLING DIRECTLY
            result_path = run_invoice_generation(
                input_data_path=json_path,
                output_path=output_path,
                template_dir=template_dir,
                config_dir=config_dir,
                daf_mode=daf_mode,
                custom_mode=custom_mode,
                explicit_config_path=explicit_config_path,
                explicit_template_path=explicit_template_path,
                input_data_dict=input_data_dict
            )
            return result_path

        except Exception as e:
            # Capture the full traceback for the UI to display
            import traceback
            tb = traceback.format_exc()
            raise RuntimeError(f"Invoice Generation Failed:\n{tb}") from e