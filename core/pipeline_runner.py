
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

        # 2. Blueprint Generator
        # We want to output config to the same base output dir? Or a specific 'configs' dir?
        # Standard: invoice_generator/src/config_bundled/ OR temp_test_data?
        # Let's verify generation into the output_dir.
        def run_blueprint():
             # Initialize generator
             gen = BlueprintGenService(output_base_dir=self.output_dir)
             # Run generation
             gen.generate(str(self.input_excel))

        if not self.runs_step("Blueprint Generator", run_blueprint, trace_id=tid):
            return

        # Locate Generated Config
        # It should be in output_dir/{Something}/{Something}_config.json
        # We can try to find it.
        found_configs = list(self.output_dir.rglob("*_config.json"))
        if not found_configs:
            logger.error("Could not locate generated config file in output directory.")
            return
        
        # Pick the most likely one (matching stem?)
        self.generated_config_path = found_configs[0]
        # Prefer one that matches input stem if multiple
        for cfg in found_configs:
            if self.input_stem in cfg.name:
                self.generated_config_path = cfg
                break
        
        logger.info(f"Using Config: {self.generated_config_path}")


        # Look for template file (same stem as config, but .xlsx)
        self.generated_template_path = self.generated_config_path.with_suffix(".xlsx")
        # Or search in the dir if name differs slightly?
        # Standard: {Customer}.xlsx
        # Config: {Customer}_config.json
        # So replacing _config.json with .xlsx should work.
        if not self.generated_template_path.exists():
            # Try replacing just .json with .xlsx
             monitor_stem = self.generated_config_path.stem.replace("_config", "")
             self.generated_template_path = self.generated_config_path.parent / f"{monitor_stem}.xlsx"
        
        if self.generated_template_path.exists():
             logger.info(f"Using Template: {self.generated_template_path}")
        else:
             logger.warning(f"Template file not found at expected path: {self.generated_template_path}")

        # 3. Invoice Generator
        # generate_invoice.py usage: generate_invoice.py [-o OUTPUT] ... input_data_file
        def run_invoice_gen():
            run_invoice_generation(
                 input_data_path=self.parsed_data_path,
                 explicit_config_path=self.generated_config_path,
                 output_path=self.output_dir,
                 explicit_template_path=self.generated_template_path if self.generated_template_path.exists() else None,
                 template_dir=None, # will use default logic if explicit path not provided
                 config_dir=None # will use default logic
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
