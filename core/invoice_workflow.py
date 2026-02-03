
import subprocess
import sys
import json
import logging
import argparse
from pathlib import Path
from typing import Optional, Dict, Any

# Initialize centralized logging FIRST before other core imports
from core.logger_config import setup_logging
from core.system_config import sys_config
setup_logging(log_dir=sys_config.run_log_dir)

logger = logging.getLogger(__name__)

from core.data_parser.main import run_invoice_automation
from core.blueprint_generator.blueprint_generator import BlueprintGenerator as BlueprintGenService
from core.invoice_generator.generate_invoice import run_invoice_generation


class InvoiceWorkflow:
    """
    Orchestrates the complete invoice processing workflow:
    1. Data Parsing (Excel -> JSON)
    2. Bundle Resolution (find matching template/config)
    3. Invoice Generation (JSON + Template -> Output Excel)
    
    Writes centralized metadata after completion.
    """
    def __init__(self, input_excel: str, output_dir: str):
        self.input_excel = Path(input_excel).resolve()
        self.output_dir = Path(output_dir).resolve()
        self.input_stem = self.input_excel.stem
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Define expected artifact paths
        self.parsed_data_path = self.output_dir / f"{self.input_stem}.json"
        self.generated_config_path = None
        self.generated_template_path = None

    def _resolve_bundle_path(self) -> Optional[Path]:
        """
        Resolve the Bundle Directory based on the input filename prefix.
        E.g., Input 'JF25057.xlsx' -> matches folder 'JF' -> returns '.../bundled/JF'
        """
        bundled_dir = sys_config.bundled_dir
        
        if not bundled_dir.exists():
            return None

        # Sort by length descending to match longest prefix first
        candidates = [d for d in bundled_dir.iterdir() if d.is_dir()]
        candidates.sort(key=lambda x: len(x.name), reverse=True)
        
        input_name_upper = self.input_stem.upper()
        
        for folder in candidates:
            if input_name_upper.startswith(folder.name.upper()):
                logger.info(f"Resolved Bundle: {folder.name} (for {self.input_stem})")
                return folder
        
        return None

    def _run_step(self, step_name: str, func, trace_id: str = 'unknown') -> bool:
        """Execute a workflow step with error handling."""
        logger.info(f"--- [{step_name}] Started [{trace_id}] ---")
        
        try:
            func()
            logger.info(f"--- [{step_name}] Completed ---")
            return True
        except Exception as e:
            logger.error(f"[{step_name}] Failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def execute(self):
        """Execute the complete invoice workflow."""
        from core.utils.snitch import start_trace
        trace_id = start_trace()
        logger.info(f"=== Invoice Workflow Started: {self.input_excel.name} | Trace: {trace_id} ===")
        
        # Step 1: Parse Data
        def parse_data():
            run_invoice_automation(
                input_excel_override=str(self.input_excel),
                output_dir_override=str(self.output_dir)
            )
             
        parser_success = self._run_step("Data Parser", parse_data, trace_id)
        if not parser_success:
            self._write_workflow_metadata(False, False, None, None)
            return

        # Step 2: Resolve Bundle
        bundle_path = self._resolve_bundle_path()
        if not bundle_path:
            logger.error(f"No matching Bundle for '{self.input_stem}'. Ensure bundle exists in database/blueprints/bundled/")
            self._write_workflow_metadata(parser_success, False, None, None)
            return

        # Derive config and template paths
        customer_code = bundle_path.name
        possible_config = bundle_path / f"{customer_code}_bundle_config.json"
        
        if not possible_config.exists():
            found_configs = list(bundle_path.glob("*_config.json"))
            if found_configs:
                possible_config = found_configs[0]
                logger.info(f"Resolved Config: {possible_config.name}")
        
        self.generated_config_path = possible_config
        self.generated_template_path = bundle_path / f"{customer_code}.xlsx"

        if not self.generated_config_path.exists():
            logger.error(f"Bundle Config not found: {self.generated_config_path}")
            self._write_workflow_metadata(parser_success, False, bundle_path, None)
            return
             
        if not self.generated_template_path.exists():
            logger.warning(f"Bundle Template not found: {self.generated_template_path}")
             
        logger.info(f"Using Config: {self.generated_config_path}")
        logger.info(f"Using Template: {self.generated_template_path}")

        # Step 3: Generate Invoice
        output_file_path = self.output_dir / f"{self.input_stem}.xlsx"
        
        def generate_invoice():
            with open(self.parsed_data_path, 'r', encoding='utf-8') as f:
                data_dict = json.load(f)

            run_invoice_generation(
                input_data_path=self.parsed_data_path,
                explicit_config_path=self.generated_config_path,
                output_path=output_file_path,
                explicit_template_path=self.generated_template_path if self.generated_template_path.exists() else None,
                template_dir=None,
                config_dir=None,
                input_data_dict=data_dict
            )
        
        invoice_success = self._run_step("Invoice Generator", generate_invoice, trace_id)
        
        # Write Centralized Metadata
        self._write_workflow_metadata(parser_success, invoice_success, bundle_path, output_file_path)
            
        logger.info("=== Invoice Workflow Completed ===")

    def _write_workflow_metadata(self, parser_success: bool, invoice_success: bool, 
                                  bundle_path: Optional[Path], output_file: Optional[Path]):
        """Write a single, centralized metadata file for the workflow."""
        import datetime
        
        meta_path = self.output_dir / f"{self.input_stem}_workflow_metadata.json"
        
        metadata = {
            "status": "success" if (parser_success and invoice_success) else "failed",
            "timestamp": datetime.datetime.now().isoformat(),
            "input_file": self.input_excel.name,
            "output_file": output_file.name if output_file and output_file.exists() else None,
            "bundle_used": bundle_path.name if bundle_path else None,
            "steps": {
                "data_parser": {
                    "success": parser_success,
                    "output": self.parsed_data_path.name if self.parsed_data_path.exists() else None
                },
                "invoice_generator": {
                    "success": invoice_success,
                    "output": output_file.name if output_file and output_file.exists() else None
                }
            }
        }
        
        try:
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=4)
            logger.info(f"Workflow metadata: {meta_path}")
        except Exception as e:
            logger.error(f"Failed to write workflow metadata: {e}")


# Backward compatibility alias
PipelineOrchestrator = InvoiceWorkflow


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run Invoice Workflow")
    parser.add_argument("input", help="Input Excel File")
    parser.add_argument("--output", default=str(sys_config.output_dir), help="Output Directory")
    
    args = parser.parse_args()
    
    workflow = InvoiceWorkflow(args.input, args.output)
    workflow.execute()
