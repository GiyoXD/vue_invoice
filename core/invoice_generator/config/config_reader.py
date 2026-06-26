import json
import logging
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


class ConfigFileReader:
    """
    Stateless reader to load and parse configuration and sibling template files from disk.
    """
    
    @staticmethod
    def load(config_path: Path) -> Tuple[Dict[str, Any], Optional[Dict[str, Any]]]:
        """
        Load a configuration JSON file and its sibling template JSON file.
        
        Args:
            config_path: Path to the configuration JSON file.
            
        Returns:
            Tuple of (config_data, template_data).
        """
        logger.debug(f"Loading configuration from: {config_path}")
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # Load sibling template config for JSON-based reconstruction
            template_data = None
            try:
                # Deduce template json path: same dir, "{config_name}_template.json"
                # Convention: {CLIENT}_config.json -> {CLIENT}_template.json
                # OR just side by side replacement: _config.json -> _template.json
                stem = config_path.stem
                parent = config_path.parent
                
                if stem.endswith('_config'):
                    template_name = stem.replace('_config', '_template') + ".json"
                else:
                    template_name = f"{stem}_template.json"
                    
                template_path = parent / template_name
                if template_path.exists():
                    with open(template_path, 'r', encoding='utf-8') as f:
                        raw_tmpl = json.load(f)
                        # The file usually has root {"template_layout": {...}}
                        template_data = raw_tmpl.get("template_layout", {})
                        logger.info(f"Loaded sibling template config from: {template_path}")
                else:
                    logger.debug(f"No sibling template JSON found at {template_path}")
            except Exception as e:
                logger.warning(f"Failed to load sibling template JSON: {e}")
                
            return config_data, template_data
            
        except Exception as e:
            logger.error(f"Error loading configuration file {config_path}: {e}")
            raise
