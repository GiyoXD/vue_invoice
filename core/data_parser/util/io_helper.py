import os
import logging
from pathlib import Path
from typing import Tuple, Any, Optional, Union
from core.data_parser import config as cfg

logger = logging.getLogger(__name__)

def resolve_output_dir(output_dir_override: Optional[str] = None) -> Path:
    """
    Determines and prepares the output directory.
    If override is specified, attempts to create it (raising RuntimeError on failure).
    Otherwise, defaults to sys_config.temp_uploads_dir.
    """
    if output_dir_override:
        output_dir = Path(output_dir_override).resolve()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise RuntimeError(f"Invalid output directory specified: {output_dir}") from e
    else:
        from core.system_config import sys_config
        output_dir = sys_config.temp_uploads_dir
    return output_dir

def prepare_input_source(
    input_excel_override: Any,
    input_filename_override: Optional[str] = None
) -> Tuple[Any, str, bool]:
    """
    Determines the input filepath/buffer, the filename, and whether it's a buffer.
    """
    is_buffer = hasattr(input_excel_override, "read")
    if is_buffer:
        input_filepath = input_excel_override
        input_name = input_filename_override or "upload.xlsx"
    else:
        input_filepath = input_excel_override or getattr(cfg, 'INPUT_EXCEL_FILE', 'unknown.xlsx')
        input_name = Path(input_filepath).name
    
    return input_filepath, input_name, is_buffer

def validate_and_resolve_filepath(
    input_filepath: Any,
    data_parser_dir: Union[str, Path],
    has_override: bool = False
) -> str:
    """
    Validates the input filepath. Resolves relative paths if necessary.
    Meant to be called inside the PipelineMonitor block to capture configuration errors.
    """
    if not has_override:
        try:
            # Verify the config has the excel file path set
            _ = cfg.INPUT_EXCEL_FILE
        except AttributeError as e:
            raise RuntimeError("Input Excel file path is missing in config.") from e

    filepath_str = str(input_filepath)

    if not os.path.isfile(filepath_str):
        # Try relative resolution
        potential_path = os.path.join(data_parser_dir, filepath_str)
        if os.path.isfile(potential_path):
            filepath_str = potential_path
            logger.info(f"Resolved relative input path: {filepath_str}")
        else:
            raise FileNotFoundError(f"Input Excel file not found: {filepath_str}")

    return filepath_str
