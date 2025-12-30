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
        Strategy 1: Look for a dedicated folder for this configuration (Registry).
        Expected structure: .../registry/{Client}/ or .../registry/{Client}_config/
        """
        # We try strict match first, then prefix match
        candidates = [file_stem]
        
        # Add prefix candidate (e.g. JF25058 -> JF, CT25048E -> CT)
        # Capture leading naming code (ignoring numbers and subsequent characters)
        match = re.match(r'^([a-zA-Z]+)', file_stem)
        prefix = match.group(1) if match else None
        
        if prefix and prefix != file_stem:
            candidates.append(prefix)

        for candidate in candidates:
            # We look for a folder named "{Candidate}_config"
            # It might be fuzzy, but strict naming is safer. 
            # Let's try to find a folder that *starts* with the candidate and ends with _config
            # directly constructing the path is faster and safer than iterating directories.
            
            # Common pattern: "{Candidate}_config" or just "{Candidate}"? 
            # Based on user data: "CT&INV&PL JF25058 FCA_config"
            
            # Since the folder name might be complex (e.g. "CT&INV&PL JF25058 FCA_config"),
            # we might need to search the directory if a direct match fails.
            
            # Direct check
            folder_name = candidate
            potential_dir = self.config_dir / folder_name
            
            if potential_dir.exists() and potential_dir.is_dir():
                return self._get_assets_from_folder(potential_dir, candidate)

            # Iterative check for partial matches (more expensive but more flexible)
            # Find any folder STARTING with the candidate (more precise than 'in')
            
            # User example: Stem="JF25058", Folder="CT&INV&PL JF25058 FCA_config" -> This case works if candidate is 'CT'?
            # Wait, if Stem is JF25058, candidate is JF. Folder "CT&INV..." does NOT start with JF.
            # But the user example "CT25048E" -> "CT" -> Folder "CT_config". THIS works.
            
            for folder in self.config_dir.iterdir():
                if folder.is_dir() and folder.name.startswith(candidate):
                     assets = self._get_assets_from_folder(folder, candidate)
                     if assets:
                         return assets
        
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
        Strategy 2: Look for standalone files in the root config directory.
        Legacy behavior support.
        """
        # 1. Look for Config
        config_path = self.config_dir / f"{file_stem}_bundle_config.json"
        effective_stem = file_stem

        if not config_path.exists():
            # Try prefix (JF25058 -> JF)
            prefix = re.sub(r'[\d_]+$', '', file_stem)
            if prefix and prefix != file_stem:
                prefix_config = self.config_dir / f"{prefix}_bundle_config.json"
                if prefix_config.exists():
                    config_path = prefix_config
                    effective_stem = prefix

        if not config_path.exists():
            # Try simple name
            config_path = self.config_dir / f"{file_stem}.json"
        
        if not config_path.exists():
            return None

        # 2. Look for Template (in the template directory, not config directory)
        # Template name derivation usually relies on config metadata, but here we guess by name first
        template_path = self.template_dir / f"{effective_stem}.xlsx"
        
        if not template_path.exists():
            # Try to read config to find template name?
            # For simplicity/speed in this step, we check generic fallback or strict name.
            # If explicit resolving fails, we might miss the metadata-defined template name.
            # Let's peek into config quickly? 
            template_path = self._peek_config_for_template_name(config_path)
            
            if not template_path:
                 template_path = self.template_dir / "Invoice.xlsx" # Generic fallback

        if template_path and template_path.exists():
            return InvoiceAssets(Path(""), config_path, template_path)
            
        return None

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
