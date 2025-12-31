"""
resolvers.py

This module provides the logic to "resolve" (find) the necessary assets for invoice generation.
"Assets" specifically refers to:
1. The Configuration File (.json) - Defines how to process the data.
2. The Template File (.xlsx) - The Excel file to be filled with data.

These assets are located based on the input data filename using a prioritized strategy:
1. Registry Strategy (formerly Bundled): Checks if a dedicated folder exists for this client aka "Registry" (e.g. database/blueprints/registry/CLIENT/).
2. Flat File Strategy: Checks for standalone files in the main config directory.
3. Fallback Strategy: Uses default assets if specific ones are not found.
"""

import os
import json
import re
import logging
from pathlib import Path
from typing import Optional, Dict, NamedTuple

logger = logging.getLogger(__name__)

class InvoiceAssets(NamedTuple):
    """Holds the resolved paths for generation assets."""
    data_path: Path
    config_path: Path
    template_path: Path

class InvoiceAssetResolver:
    """
    Responsible for locating the Configuration and Template files required to generate an invoice.
    It hides the complexity of where these files are stored (flat lists vs bundled folders).
    """

    def __init__(self, base_config_dir: Path, base_template_dir: Path):
        self.config_dir = Path(base_config_dir)
        self.template_dir = Path(base_template_dir)

    def resolve_assets_for_input_file(self, input_file_path: str) -> Optional[InvoiceAssets]:
        """
        Main entry point. Finds the config and template needed to process the given input file.
        
        Args:
            input_file_path: The path to the user's input data file (e.g. 'JF25058.json')
            
        Returns:
            InvoiceAssets object containing all three paths, or None if assets generally failed to resolve.
        """
        input_path = Path(input_file_path)
        stem = input_path.stem
        
        logger.info(f"Resolving assets for input: {stem}")

        # 1. Try to find a "Registry" configuration (Self-contained folder)
        #    Example: registry/JF/
        registry_assets = self._try_resolve_from_registry(stem)
        if registry_assets:
            logger.info(f"✅ Resolved assets using Registry Strategy: {registry_assets.config_path.parent.name}")
            return InvoiceAssets(input_path, registry_assets.config_path, registry_assets.template_path)

        # 2. Try to find "Flat" configuration (Standalone files)
        #    Example: config/bundled/JF25058_bundle_config.json
        flat_assets = self._try_resolve_flat_files(stem)
        if flat_assets:
            logger.info(f"✅ Resolved assets using Flat File Strategy: {flat_assets.config_path.name}")
            return InvoiceAssets(input_path, flat_assets.config_path, flat_assets.template_path)

        # 3. Default Fallback -> REMOVED
        #    If nothing specific found, we fail. We do NOT fallback to default.json anymore.
        
        # fallback_assets = self._resolve_fallback()
        # if fallback_assets:
        #     logger.warning(f"⚠️ specific config not found. Using Fallback Strategy: {fallback_assets.config_path.name}")
        #     return InvoiceAssets(input_path, fallback_assets.config_path, fallback_assets.template_path)

        logger.error(f"❌ Could not resolve any valid assets for {stem}")
        return None

    def _try_resolve_from_registry(self, file_stem: str) -> Optional[InvoiceAssets]:
        """
        Strategy 1: Look for a bundled folder using PREFIX matching only.
        
        Input: JF25061 → Look for bundled/JF/ (extract letters prefix)
        Input: CT25048E → Look for bundled/CT/
        
        We NEVER check for the full filename (e.g., bundled/JF25061/) because
        those folders will never exist - only the prefix-based ones do.
        """
        # Extract prefix (letters only, e.g., JF25058 -> JF, CT25048E -> CT)
        match = re.match(r'^([a-zA-Z]+)', file_stem)
        prefix = match.group(1) if match else None
        
        if not prefix:
            logger.warning(f"Could not extract prefix from '{file_stem}'")
            return None
        
        logger.info(f"Looking for bundle using prefix: '{prefix}' (from '{file_stem}')")
        
        # Direct check for prefix folder
        potential_dir = self.config_dir / prefix
        
        if potential_dir.exists() and potential_dir.is_dir():
            return self._get_assets_from_folder(potential_dir, prefix)
        
        # Fallback: Check for folders starting with prefix (e.g., JF_config, JF_v2)
        # Only iterate if the config directory exists
        if self.config_dir.exists() and self.config_dir.is_dir():
            for folder in self.config_dir.iterdir():
                if folder.is_dir() and folder.name.startswith(prefix):
                    assets = self._get_assets_from_folder(folder, prefix)
                    if assets:
                        return assets
        
        logger.warning(f"No bundle folder found for prefix '{prefix}' in {self.config_dir}")
        return None

    def _get_assets_from_folder(self, folder_path: Path, identifier: str) -> Optional[InvoiceAssets]:
        """Helper to extract config and template from a specific valid folder."""
        # Inside the folder, we look for key files.
        # 1. Config: Ends with .json
        # 2. Template: Ends with .xlsx (excluding temporary ones)
        
        config_file = None
        template_file = None
        
        # Heuristic: Find the "main" config file. 
        # Usually named same as folder or "{Identifier}_config.json"
        for f in folder_path.iterdir():
            if f.suffix.lower() == '.json':
                # Avoid "template_config.json" if possible, usually we want the main bundle config
                if "_template" not in f.name:
                    config_file = f
            elif f.suffix.lower() == '.xlsx':
                if not f.name.startswith('~$'): # Ignore temp lock files
                    template_file = f
        
        if config_file and template_file:
            # We found a pair!
            return InvoiceAssets(Path(""), config_file, template_file) # Data path is dummy here, rewritten later
        
        return None

    def _try_resolve_flat_files(self, file_stem: str) -> Optional[InvoiceAssets]:
        """
        Strategy 2 (Legacy): Look for flat config files using PREFIX only.
        
        This is a fallback for configs not in bundled folders.
        Input: JF25061 → Look for JF_bundle_config.json or JF_config.json
        """
        # Extract prefix (letters only)
        match = re.match(r'^([a-zA-Z]+)', file_stem)
        prefix = match.group(1) if match else None
        
        if not prefix:
            return None
        
        # Check for prefix-based config files
        config_candidates = [
            self.config_dir / f"{prefix}_bundle_config.json",
            self.config_dir / f"{prefix}_config.json",
            self.config_dir / f"{prefix}.json"
        ]
        
        config_path = None
        for candidate in config_candidates:
            if candidate.exists():
                config_path = candidate
                break
        
        if not config_path:
            return None

        # Look for template in same directory (bundled approach)
        template_path = self.template_dir / f"{prefix}.xlsx"
        
        if not template_path.exists():
            # Try to read config to find template name
            template_path = self._peek_config_for_template_name(config_path)
            
        if not template_path or not template_path.exists():
            return None

        return InvoiceAssets(Path(""), config_path, template_path)

    def _peek_config_for_template_name(self, config_path: Path) -> Optional[Path]:
        """Reads the _meta section of a config file to find the linked template name."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                template_name = data.get('_meta', {}).get('template_name')
                if template_name:
                    return self.template_dir / template_name
        except Exception:
            pass
        return None
