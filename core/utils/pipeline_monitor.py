
import logging
import time
import datetime
import json
import traceback
import sys
from pathlib import Path
from typing import Optional, List, Any, Dict, Union

logger = logging.getLogger(__name__)

class PipelineMonitor:
    """
    Context manager to monitor pipeline execution, track state, and GUARANTEE 
    metadata file generation upon exit (success or failure).
    
    Unified version for Data Parser, Blueprint Generator, and Invoice Generator.
    """
    def __init__(self, output_path: Union[str, Path], args: Any = None, input_data: Dict = None, step_name: str = "unknown"):
        self.output_path = Path(output_path)
        self.args = args
        self.input_data = input_data or {}
        self.step_name = step_name
        
        self.start_time = None
        self.duration = 0.0
        
        # Tracking Lists
        self.items_processed = [] # Generic "items" (sheets, tables, etc.)
        self.items_failed = []
        self.warnings = [] # List of warning strings
        
        self.logs = {} # Generic dictionary for step-specific logs (replacements, stats)
        
        self.status = "pending"
        self.error_message = None
        self.error_traceback = None
        self.exit_code = 0

    def __enter__(self):
        self.start_time = time.time()
        logger.info(f"=== [{self.step_name}] Process Started ===")
        return self

    def log_process_item(self, item_name: str, status: str = "success", error: Exception = None):
        """Log the processing of a specific sub-item (e.g., a sheet or table)."""
        if status == "success":
            self.items_processed.append(item_name)
            logger.info(f"[{self.step_name}] Successfully processed: {item_name}")
        else:
            self.items_failed.append(item_name)
            msg = f"[{self.step_name}] Failed to process {item_name}: {error}"
            self.warnings.append(msg)
            logger.error(msg)
            if error:
                logger.debug(traceback.format_exc())

    def log_warning(self, message: str):
        """Explicitly log a warning that didn't stop execution but should be noted."""
        self.warnings.append(message)
        logger.warning(f"[{self.step_name}] WARNING: {message}")

    def update_logs(self, key: str, data: Any):
        """Update generic step-specific logs."""
        if key not in self.logs:
            self.logs[key] = []
        
        if isinstance(self.logs[key], list) and isinstance(data, list):
            self.logs[key].extend(data)
        elif isinstance(self.logs[key], dict) and isinstance(data, dict):
            self.logs[key].update(data)
        else:
            self.logs[key] = data

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.start_time:
            self.duration = time.time() - self.start_time
        
        # Determine status
        if exc_type:
            self.status = "fatal"
            self.error_message = str(exc_val)
            self.error_traceback = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
            self.exit_code = 1
            logger.critical(f"[{self.step_name}] Process Crashed: {self.error_message}")
        elif self.items_failed:
            self.status = "partial_success" if self.items_processed else "failure"
            self.error_message = f"Failed items: {self.items_failed}"
            # We might want exit_code 0 for partial success, or 1? 
            # Usually strict pipelines prefer 1 if *any* failure.
            # For now, keep 0 unless fatal, but status reflects issues.
        elif self.warnings:
            self.status = "success_with_warnings"
        else:
            self.status = "success"

        # Generate Metadata
        try:
            self._write_metadata()
        except Exception as e:
            logger.error(f"[{self.step_name}] Failed to write metadata: {e}")
            
        # We generally propagate fatal exceptions so CLI knows to crash
        return False 

    def _write_metadata(self):
        """Write the metadata JSON file."""
        # [DISABLED] Per user request, we do not want these JSON files.
        return

        # Ensure output directory exists
        if not self.output_path.parent.exists():
            try:
                self.output_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                logger.error(f"Failed to create output directory for metadata: {e}")

        # Metadata Filename: [original_name]_metadata.json
        meta_filename = f"{self.output_path.stem}_metadata.json"
        
        # Logic to handle if output_path IS the metadata file itself or the result file
        if self.output_path.suffix == ".json" and "_metadata" in self.output_path.name:
            meta_path = self.output_path
        else:
             meta_path = self.output_path.parent / meta_filename
        
        # Construct Metadata
        metadata = {
            "step": self.step_name,
            "status": self.status,
            "timestamp": datetime.datetime.now().isoformat(),
            "duration_seconds": self.duration,
            "exit_code": self.exit_code,
            "items_processed": self.items_processed,
            "items_failed": self.items_failed,
            "warnings": self.warnings,
            "error_message": self.error_message,
            "error_traceback": self.error_traceback,
            "custom_logs": self.logs
        }
        
        # Add Input/Config info if available
        if self.args:
             try:
                 metadata["args"] = vars(self.args)
             except:
                 metadata["args"] = str(self.args)

        with open(meta_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=4)
        
        logger.info(f"[{self.step_name}] Metadata written to {meta_path}")

