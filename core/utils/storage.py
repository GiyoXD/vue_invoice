import shutil
import logging
from pathlib import Path
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
    def clear_runtime_cache(customer_code: str, locale: str) -> None:
        """Clean up specific runtime cache path for a customer/locale variant."""
        temp_dir = sys_config.temp_uploads_dir / "runtime_blueprints" / f"{customer_code}_{locale}"
        TempFileStorage.clear_directory(temp_dir)
