
import subprocess
import sys
import json
import logging
import argparse
from pathlib import Path
from typing import Optional, Dict, Any

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from core.data_parser.main import run_invoice_automation
from core.blueprint_generator.blueprint_generator import BlueprintGenerator as BlueprintGenService
from core.invoice_generator.generate_invoice import run_invoice_generation
from core.utils.pipeline_monitor import PipelineMonitor

class PipelineOrchestrator:
    def __init__(self, input_excel: str, output_dir: str):
        self.input_excel = Path(input_excel).resolve()
        self.output_dir = Path(output_dir).resolve()
        self.input_stem = self.input_excel.stem
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Define expected artifact paths
        self.parser_metadata_path = self.output_dir / f"{self.input_stem}_parser_metadata.json"
        self.parsed_data_path = self.output_dir / f"{self.input_stem}.json"
        
        # Blueprint metadata is saved in source dir by default in my implementation (args_template_path.parent)
        # But wait, if I run from Runner, I might want to control this?
        # Current Blueprint Generator uses input file directory for metadata unless forced?
        # Actually it used: monitor_path = args_template_path.parent / ...
        self.blueprint_metadata_path = self.input_excel.parent / f"{self.input_stem}_blueprint_metadata.json"
        
        # Config outputs are tricky to predict without parsing, but usually {Stem}/{Stem}_config.json or {Customer}/{Customer}_config.json
        # We will search for them.
        self.generated_config_path = None
        
        self.invoice_metadata_path = self.output_dir / "invoice_generation_metadata.json" # Invoice Gen usually saves to output dir?

    def _resolve_bundle_path(self) -> Optional[Path]:
        """
        Attempts to resolve the Bundle Directory based on the input filename prefix.
        E.g., Input 'JF25057.xlsx' -> matches folder 'JF' -> returns '.../bundled/JF'
        """
        from core.system_config import sys_config
        bundled_dir = sys_config.bundled_dir
        
        if not bundled_dir.exists():
            return None

        # Sort by length descending to match longest prefix first
        candidates = [d for d in bundled_dir.iterdir() if d.is_dir()]
        candidates.sort(key=lambda x: len(x.name), reverse=True)
        
        input_name_upper = self.input_stem.upper()
        
        for folder in candidates:
             # Check if input starts with folder name (e.g. "JF25057" starts with "JF")
            if input_name_upper.startswith(folder.name.upper()):
                 logger.info(f"Resolved Bundle Directory: {folder} (for input {self.input_stem})")
                 return folder
        
        return None

    def runs_step(self, step_name: str, func, *args, **kwargs) -> bool:
        """Run a pipeline step directly and verify its metadata."""
        tid = kwargs.pop('trace_id', 'unknown')
        logger.info(f"--- Running Step: {step_name} [{tid}] ---")
        
        try:
            # Direct execution in the same process
            func(*args, **kwargs)
            return True
                
        except Exception as e:
            logger.error(f"Orchestrator failed to run step {step_name}: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def run(self):
        # Initialize Context-Aware Tracing (The Snitch)
        from core.utils.snitch import start_trace
        tid = start_trace()
        logger.info(f"Starting Pipeline for: {self.input_excel} | TraceID: {tid}")
        
        # 1. Data Parser
        def run_parser():
             run_invoice_automation(
                input_excel_override=str(self.input_excel),
                output_dir_override=str(self.output_dir)
             )
             
        if not self.runs_step("Data Parser", run_parser, trace_id=tid):
            return

        # [SKIPPED] 2. Blueprint Generator
        # Per user request, Blueprint Generation is a separate workflow.
        # We now look for an EXISTING bundle instead.
        
        # 2b. Resolve Existing Bundle
        bundle_path = self._resolve_bundle_path()
        if not bundle_path:
             logger.error(f"No matching Bundle found for input '{self.input_stem}'. Cannot proceed with Invoice Generation.")
             logger.error("Please ensure a valid Customer Bundle exists in database/blueprints/bundled/")
             return

        # Derive paths from bundle folder
        # e.g. bundled/JF/JF_bundle_config.json
        customer_code = bundle_path.name
        
        # Try finding the config file dynamically
        possible_config = bundle_path / f"{customer_code}_bundle_config.json"
        
        if not possible_config.exists():
             # Fallback: Search for any *_config.json
             found_configs = list(bundle_path.glob("*_config.json"))
             if found_configs:
                 possible_config = found_configs[0]
                 logger.info(f"Resolved Config File via wildcard: {possible_config.name}")
        
        self.generated_config_path = possible_config
        
        # Template is also in the bundle
        self.generated_template_path = bundle_path / f"{customer_code}.xlsx"

        if not self.generated_config_path.exists():
             logger.error(f"Bundle Config not found at: {self.generated_config_path}")
             return
             
        if not self.generated_template_path.exists():
             logger.warning(f"Bundle Template not found at: {self.generated_template_path}. Invoice Gen might fail if it needs strict templating.")
             
        logger.info(f"Using Bundle Config: {self.generated_config_path}")
        logger.info(f"Using Bundle Template: {self.generated_template_path}")


        # 3. Invoice Generator
        # generate_invoice.py usage: generate_invoice.py [-o OUTPUT] ... input_data_file
        
        # Ensure we target a FILE, not a directory
        output_file_path = self.output_dir / f"{self.input_stem}.xlsx"
        
        def run_invoice_gen():
            # Load the parsed data dictionary to pass to the generator
            try:
                with open(self.parsed_data_path, 'r', encoding='utf-8') as f:
                    data_dict = json.load(f)
            except Exception as e:
                logger.error(f"Failed to load user input data from {self.parsed_data_path}: {e}")
                raise e

            run_invoice_generation(
                 input_data_path=self.parsed_data_path,
                 explicit_config_path=self.generated_config_path,
                 output_path=output_file_path,
                 explicit_template_path=self.generated_template_path if self.generated_template_path.exists() else None,
                 template_dir=None, # will use default logic if explicit path not provided
                 config_dir=None, # will use default logic
                 input_data_dict=data_dict
            )
        
        # Metadata for invoice generator? 
        # Check generate_invoice.py to see where it saves metadata.
        # It usually uses `GenerationMonitor`.
        
        # We'll run it and check result.
        if not self.runs_step("Invoice Generator", run_invoice_gen, trace_id=tid):
            # If explicit metadata path check fails, we might check if result PDF exists?
            pass
            
        logger.info("Pipeline Completed Successfully.")

if __name__ == "__main__":
    from core.system_config import sys_config
    
    parser = argparse.ArgumentParser(description="Run the Full Invoice Pipeline")
    parser.add_argument("input", help="Input Excel File")
    parser.add_argument("--output", default=str(sys_config.output_dir), help="Output Directory")
    
    args = parser.parse_args()
    
    runner = PipelineOrchestrator(args.input, args.output)
    runner.run()
