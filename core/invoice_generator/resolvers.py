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
from typing import Optional, Dict, NamedTuple, Any
from core.system_config import sys_config

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
    Supports KH/VN variant detection within bundle folders.
    """

    VARIANT_SUFFIXES = ["_KH", "_VN"]

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
        """
        Helper to extract config and template from a specific valid folder.
        
        If the folder contains KH/VN variants, returns the first variant found
        as the primary asset (so existing single-config flow still works).
        """
        # Check for KH/VN variants first
        variants = self.resolve_variants_from_folder(folder_path, identifier)
        if variants:
            # Return first variant as the primary (backward compatible)
            first = variants[0]
            return InvoiceAssets(Path(""), first["config_path"], first["template_path"])
        
        # Fallback: look for a single config/template pair (original logic)
        config_file = None
        template_file = None
        
        for f in folder_path.iterdir():
            if f.suffix.lower() == '.json':
                if "_template" not in f.name:
                    config_file = f
            elif f.suffix.lower() == '.xlsx':
                if not f.name.startswith('~$'):
                    template_file = f
        
        if config_file and template_file:
            return InvoiceAssets(Path(""), config_file, template_file)
        
        return None

    def resolve_variants_from_folder(self, folder_path: Path, prefix: str):
        """
        Detect KH/VN variant configs and templates in a bundle folder.
        
        Looks for files matching pattern:
          {prefix}{suffix}_config.json + {prefix}{suffix}.xlsx
        
        Args:
            folder_path: Path to the bundle folder (e.g. bundled/TEST/)
            prefix: The base prefix (e.g. 'TEST')
            
        Returns:
            List of dicts: [{suffix, config_path, template_path}, ...]
            Empty list if no variants found.
        """
        variants = []
        for suffix in self.VARIANT_SUFFIXES:
            config = folder_path / f"{prefix}{suffix}_config.json"
            template = folder_path / f"{prefix}{suffix}.xlsx"
            if config.exists() and template.exists():
                variants.append({
                    "suffix": suffix,
                    "config_path": config,
                    "template_path": template
                })
        return variants

    def resolve_all_variants(self, input_file_path: str):
        """
        Public method to find all KH/VN variants for an input file.
        
        Args:
            input_file_path: Path to the input data file (e.g. 'TEST25001.json')
            
        Returns:
            List of variant dicts, or empty list if no variants found.
        """
        input_path = Path(input_file_path)
        stem = input_path.stem
        
        match = re.match(r'^([a-zA-Z]+)', stem)
        prefix = match.group(1) if match else None
        if not prefix:
            return []
        
        potential_dir = self.config_dir / prefix
        if potential_dir.exists() and potential_dir.is_dir():
            return self.resolve_variants_from_folder(potential_dir, prefix)
        
        return []

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

# --- Legacy Helper Function ---
def derive_paths(input_data_path: str, template_dir: str, config_dir: str) -> Optional[Dict[str, Path]]:
    """
    Derive paths for config and template based on input data filename.
    Moved from generate_invoice.py.
    """
    input_path = Path(input_data_path)
    stem = input_path.stem
    
    # Prioritize bundle config to avoid picking up data file as config
    config_path = Path(config_dir) / f"{stem}_bundle_config.json"
    
    # Heuristic: If exact match not found, try stripping trailing numbers/underscores (e.g., JF25057 -> JF)
    effective_stem = stem
    if not config_path.exists():
        prefix = re.sub(r'[\d_]+$', '', stem)
        if prefix and prefix != stem:
            prefix_config = Path(config_dir) / f"{prefix}_bundle_config.json"
            if prefix_config.exists():
                config_path = prefix_config
                effective_stem = prefix # Use the prefix for template lookup too
                logger.info(f"Found config using prefix match: '{stem}' -> '{prefix}'")

    if not config_path.exists():
        config_path = Path(config_dir) / f"{stem}.json"
    
    # Fallback to default config if specific not found
    if not config_path.exists():
        default_config = Path(config_dir) / "default.json"
        if default_config.exists():
             config_path = default_config
        else:
             # If no config found, we can't proceed unless we have a strategy
             pass

    # Template path - ideally derived from config, but we need config first.
    # Strategy: Load config, check for template name. If not, use stem.
    template_path = None
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                # Check if template is specified in meta or processing
                template_name = cfg.get('_meta', {}).get('template_name')
                if template_name:
                    template_path = Path(template_dir) / template_name
        except:
            pass
    
    if not template_path:
        # Try effective stem first (e.g. JF.xlsx)
        template_path = Path(template_dir) / f"{effective_stem}.xlsx"
        if not template_path.exists() and effective_stem != stem:
             # Try original stem if effective stem failed (e.g. JF25057.xlsx)
             template_path = Path(template_dir) / f"{stem}.xlsx"

        if not template_path.exists():
             # Fallback to generic Invoice.xlsx or configured default (JF.xlsx)
             fallback = Path(template_dir) / sys_config.default_template_name
             if fallback.exists():
                 template_path = fallback

    if config_path.exists() and template_path and template_path.exists():
        return {
            'data': input_path,
            'config': config_path,
            'template': template_path
        }
    
    logger.error(f"Could not derive paths. Config: {config_path} (Exists: {config_path.exists()}), Template: {template_path} (Exists: {template_path and template_path.exists()})")
    return None
