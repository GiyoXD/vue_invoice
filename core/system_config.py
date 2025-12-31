import json
import logging
import os
from pathlib import Path
from typing import Dict, Any

# Define Project Root (Assuming this file is in core/system_config.py)
# Define Project Root (Assuming this file is in core/system_config.py)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

logger = logging.getLogger(__name__)

class SystemConfig:
    _instance = None
    # Removed _config dictionary as we no longer use system_config.json

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(SystemConfig, cls).__new__(cls)
            cls._instance._load_env_file() # Load .env variables
        return cls._instance

    def _load_env_file(self):
        """Manually load .env file into os.environ if not already set."""
        env_path = PROJECT_ROOT / ".env"
        if env_path.exists():
            try:
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith("#"):
                            continue
                        if "=" in line:
                            key, value = line.split("=", 1)
                            key = key.strip()
                            value = value.strip().strip("'").strip('"')
                            # Only set if not already in environment (OS env var takes precedence)
                            if key and key not in os.environ:
                                os.environ[key] = value
                logger.info("Loaded .env file for environment configuration.")
            except Exception as e:
                logger.warning(f"Failed to parse .env file: {e}")

    @property
    def blueprints_root(self) -> Path:
        return self._resolve_path("blueprints_root", "database/blueprints", env_key="BLUEPRINTS_ROOT")

    @property
    def bundled_dir(self) -> Path:
        """
        The bundled directory contains customer folders with config + template co-located.
        Structure: bundled/{CustomerCode}/{CustomerCode}_config.json + {CustomerCode}.xlsx
        """
        return self._resolve_path("bundled", "database/blueprints/bundled", env_key="BUNDLED_DIR")

    @property
    def templates_dir(self) -> Path:
        """For backward compatibility - points to bundled_dir since templates are now co-located."""
        return self.bundled_dir

    @property
    def registry_dir(self) -> Path:
        """For backward compatibility - points to bundled_dir since configs are now co-located."""
        return self.bundled_dir

    @property
    def mapping_config_path(self) -> Path:
        # This one is a file path, not directory usually, but logic is same
        return self._resolve_path("mapping_config", "database/blueprints/mapper/mapping_config.json", env_key="MAPPING_CONFIG")

    @property
    def output_dir(self) -> Path:
        return self._resolve_path("output", "output/generated_invoices", env_key="OUTPUT_DIR")
    
    @property
    def temp_uploads_dir(self) -> Path:
        return self._resolve_path("temp_uploads", "output/temp_uploads", env_key="TEMP_UPLOADS_DIR")
        
    @property
    def run_log_dir(self) -> Path:
        return self._resolve_path("run_log", "run_log", env_key="RUN_LOG_DIR")

    @property
    def frontend_dir(self) -> Path:
        return self._resolve_path("frontend", "frontend", env_key="FRONTEND_DIR")

    @property
    def default_template_name(self) -> str:
        env_val = os.getenv("FALLBACK_TEMPLATE")
        if env_val:
            return env_val
        return "JF.xlsx"

    def _resolve_path(self, key: str, default_relative: str, env_key: str = None) -> Path:
        # 1. Check Environmental Variable (Preferred)
        # Use env_key if provided, else uppercase key
        check_env = env_key if env_key else key.upper()
        env_val = os.getenv(check_env)
        
        if env_val:
             # Env paths are usually absolute, but if relative, assume from cwd/root
            path_obj = Path(env_val)
            if path_obj.is_absolute():
                return path_obj.resolve()
            return (PROJECT_ROOT / path_obj).resolve()

        # 2. Use Fallback Default (Hardcoded in property)
        return (PROJECT_ROOT / default_relative).resolve()

# Singleton instance
sys_config = SystemConfig()
