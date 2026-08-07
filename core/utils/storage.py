import os
import shutil
import logging
from pathlib import Path
from contextlib import contextmanager
from typing import Iterable, Union
from core.system_config import sys_config

logger = logging.getLogger(__name__)

class TempFileStorage:
    """Helper class to abstract file saving, deletion, and directory cleanup operations."""
    
    @staticmethod
    def get_temp_path(filename: str, prefix: str = "") -> Path:
        """Sanitize filename and return path in temp directory."""
        temp_dir = sys_config.temp_uploads_dir
        temp_dir.mkdir(parents=True, exist_ok=True)
        safe_name = Path(filename).name
        if prefix:
            return temp_dir / f"{prefix}_{safe_name}"
        return temp_dir / safe_name

    @staticmethod
    def save_file(file_stream, destination_path: Path) -> None:
        """Write file stream to destination path."""
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        with open(destination_path, "wb") as buffer:
            shutil.copyfileobj(file_stream, buffer)

    @staticmethod
    def delete_file(file_path: Path) -> None:
        """Delete file safely if it is a file."""
        try:
            if file_path.is_file():
                file_path.unlink()
        except Exception as e:
            logger.warning(f"Failed to delete file {file_path}: {e}")

    @staticmethod
    def clear_directory(directory_path: Path) -> None:
        """Remove a directory and all its contents recursively."""
        try:
            if directory_path.is_dir():
                shutil.rmtree(directory_path)
        except Exception as e:
            logger.warning(f"Failed to clear directory {directory_path}: {e}")

    @staticmethod
    def cleanup_customer_templates(customer_code: str, locale: str) -> None:
        """Sanitize inputs and clean up customer template directory."""
        safe_customer = os.path.basename(customer_code)
        safe_locale = os.path.basename(locale)
        temp_dir = sys_config.temp_uploads_dir / "runtime_blueprints" / f"{safe_customer}_{safe_locale}"
        TempFileStorage.clear_directory(temp_dir)

    @staticmethod
    def clear_runtime_cache(customer_code: str, locale: str) -> None:
        """Clean up specific runtime cache path for a customer/locale variant."""
        TempFileStorage.cleanup_customer_templates(customer_code, locale)


@contextmanager
def cleanup_on_failure(file_paths: Iterable[Union[str, Path]]):
    """Context manager to automatically delete temporary files if an exception occurs."""
    file_paths = list(file_paths)
    try:
        yield
    except Exception:
        for path in file_paths:
            TempFileStorage.delete_file(Path(path))
        raise


safe_temp_file_scope = cleanup_on_failure

