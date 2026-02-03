
import logging
import time
import datetime
import json
import traceback
from pathlib import Path
from typing import Optional, List, Any, Dict

logger = logging.getLogger(__name__)


class GenerationSession:
    """
    Context manager to track an invoice generation session.
    Manages lifecycle, tracks processed sheets, and logs errors.
    
    Metadata writing is centralized at the InvoiceWorkflow level.
    """
    def __init__(self, output_path: Path, args: Any = None, input_data: Dict = None):
        self.output_path = Path(output_path)
        self.args = args
        self.input_data = input_data or {}
        
        self.start_time = None
        self.sheets_processed = []
        self.sheets_failed = []
        self.replacements_log = []
        self.header_info = {}
        
        self.status = "pending"
        self.error_message = None
        self.error_traceback = None

    def __enter__(self):
        self.start_time = time.time()
        logger.info(f"=== Generation Session Started ===")
        return self

    def log_success(self, sheet_name: str):
        """Log successful processing of a sheet."""
        self.sheets_processed.append(sheet_name)
        logger.info(f"Successfully processed sheet: {sheet_name}")

    def log_failure(self, sheet_name: str, error: Exception = None):
        """Log failed processing of a sheet."""
        self.sheets_failed.append(sheet_name)
        msg = f"Failed to process sheet {sheet_name}: {error}"
        logger.error(msg)
        if error:
            logger.debug(traceback.format_exc())

    def update_logs(self, replacements: List = None, header_info: Dict = None):
        """Update session logs with replacement and header info."""
        if replacements:
            self.replacements_log.extend(replacements)
        if header_info:
            self.header_info.update(header_info)

    def __exit__(self, exc_type, exc_val, exc_tb):
        duration = time.time() - self.start_time
        
        # Determine final status
        if exc_type:
            self.status = "fatal"
            self.error_message = str(exc_val)
            self.error_traceback = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
            logger.critical(f"Session crashed: {self.error_message}")
        elif self.sheets_failed:
            self.status = "partial_success" if self.sheets_processed else "error"
            self.error_message = f"Failed sheets: {self.sheets_failed}"
        else:
            self.status = "success"

        logger.info(f"=== Generation Session Ended ({self.status}) | Duration: {duration:.2f}s ===")
        
        # Propagate exceptions
        return False

    def get_summary(self) -> Dict:
        """Return session summary for centralized metadata."""
        return {
            "status": self.status,
            "sheets_processed": self.sheets_processed,
            "sheets_failed": self.sheets_failed,
            "error_message": self.error_message
        }


# Backward compatibility alias
GenerationMonitor = GenerationSession
