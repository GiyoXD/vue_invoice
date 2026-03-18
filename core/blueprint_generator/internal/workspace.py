import hashlib
from datetime import datetime
from pathlib import Path
import shutil
import tempfile
import logging

logger = logging.getLogger(__name__)

class WorkspaceManager:
    """
    Manages file system operations, output directories, and temporary files.
    """

    def __init__(self, base_result_dir: Path):
        self.base_result_dir = base_result_dir
        self.temp_files = []

    def setup_output_directory(self, excel_file_path: Path) -> tuple[Path, dict]:
        """
        Validate the proposed directory name and enhance it to prevent collisions.
        Also detects similar existing directories that could cause confusion.

        Args:
            excel_file_path: Path to the Excel file being processed

        Returns:
            tuple: (enhanced_output_dir, metadata_dict)
        """
        original_stem = excel_file_path.stem
        self.base_result_dir.mkdir(parents=True, exist_ok=True)

        # Get existing directories to check for similar names
        existing_dirs = []
        if self.base_result_dir.exists():
            existing_dirs = [d.name for d in self.base_result_dir.iterdir() if d.is_dir()]

        # Check for potentially confusing similar names
        similar_names = []
        normalized_stem = original_stem.lower().replace('-', '').replace('_', '').replace(' ', '')

        for existing_dir in existing_dirs:
            normalized_existing = existing_dir.lower().replace('-', '').replace('_', '').replace(' ', '')
            # Check if names are very similar (differ only by punctuation)
            if normalized_stem == normalized_existing and original_stem != existing_dir:
                similar_names.append(existing_dir)

        # Create enhanced directory name with timestamp to prevent collisions
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        hash_suffix = hashlib.md5(original_stem.encode()).hexdigest()[:6]
        enhanced_dir_name = f"{original_stem}_{timestamp}_{hash_suffix}"

        enhanced_output_dir = self.base_result_dir / enhanced_dir_name
        enhanced_output_dir.mkdir(parents=True, exist_ok=True)

        # Create metadata for tracking
        metadata = {
            "original_filename": excel_file_path.name,
            "original_stem": original_stem,
            "enhanced_directory": enhanced_dir_name,
            "timestamp": datetime.now().isoformat(),
            "similar_names_detected": similar_names,
            "collision_prevention": {
                "timestamp_added": timestamp,
                "hash_suffix": hash_suffix,
                "reason": "Enhanced naming to prevent directory collisions and confusion"
            }
        }

        if similar_names:
            logger.warning("Similar directory names detected!")
            logger.warning(f"   Current file: '{original_stem}'")
            logger.warning(f"   Similar existing: {', '.join(similar_names)}")
            logger.warning("   Using enhanced directory name to prevent confusion.")

        return enhanced_output_dir, metadata

    def get_temp_file(self, suffix=".json", prefix="tmp_", dir=None) -> str:
        """Creates a temporary file and tracks it for cleanup."""
        temp_file = tempfile.NamedTemporaryFile(
            mode='w',
            delete=False,
            suffix=suffix,
            prefix=prefix,
            dir=str(dir) if dir else None
        ).name
        self.temp_files.append(temp_file)
        return temp_file

    def cleanup(self):
        """Removes tracked temporary files."""
        for file_path in self.temp_files:
            try:
                p = Path(file_path)
                if p.exists():
                    p.unlink()
                    logger.debug(f"Cleaned up temp file: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to cleanup temp file {file_path}: {e}")
        self.temp_files = []
