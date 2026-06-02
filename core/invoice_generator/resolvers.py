"""
resolvers.py

This module provides the logic to "resolve" (find) the necessary assets for invoice generation.
"Assets" specifically refers to:
1. The Configuration File (.json) - Defines how to process the data.
2. The Template File (.xlsx) - The Excel file to be filled with data.

These assets are resolved directly from the SQLite database.
"""

import json
import re
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)

@dataclass
class InvoiceAssets:
    """Holds the resolved paths for generation assets."""
    data_path: Path
    config_path: Path
    template_path: Path
    # Direct data from DB (bypasses file I/O when set)
    config_data: Optional[Dict[str, Any]] = None
    template_json_data: Optional[Dict[str, Any]] = None
    template_xlsx_bytes: Optional[bytes] = None

class InvoiceAssetResolver:
    """
    Responsible for locating the Configuration and Template files required to generate an invoice.
    Resolves directly from the database using a BlueprintRepository.
    """

    VARIANT_SUFFIXES = ["_KH", "_VN"]

    def __init__(self, base_config_dir: Path = None, base_template_dir: Path = None, repository=None):
        # We store these only for backwards-compatibility of constructor signature
        self.config_dir = Path(base_config_dir) if base_config_dir else None
        self.template_dir = Path(base_template_dir) if base_template_dir else None
        self._repository = repository

    def _resolve_from_db_row(self, row) -> Optional[InvoiceAssets]:
        """Build InvoiceAssets directly from a DB row without writing temp files."""
        try:
            config_data = json.loads(row.config_json)
            template_json_data = json.loads(row.template_json)
            xlsx_bytes = row.template_binary.xlsx_blob if row.template_binary else None
            
            return InvoiceAssets(
                data_path=Path(""),
                config_path=Path(""),  # Not used for DB blueprints
                template_path=Path(""),  # Not used for DB blueprints
                config_data=config_data,
                template_json_data=template_json_data,
                template_xlsx_bytes=xlsx_bytes
            )
        except Exception as e:
            logger.error(f"Failed to load DB blueprint for {row.customer_code}_{row.locale}: {e}")
            return None

    def _try_resolve_from_db(self, file_stem: str) -> Optional[InvoiceAssets]:
        """Database Strategy: Look for blueprint files in the database."""
        # Extract prefix
        match = re.match(r'^([a-zA-Z\-_]+)', file_stem)
        prefix = match.group(1) if match else None
        if not prefix:
            return None
        prefix = prefix.rstrip('_')
        
        # Check if the filename explicitly specifies a variant suffix
        locale = "KH"
        for suffix in ["_KH", "_VN"]:
            if suffix in file_stem.upper():
                locale = suffix.lstrip('_')
                break
        
        if self._repository:
            row = self._repository.get_blueprint(prefix, locale)
            if row:
                return self._resolve_from_db_row(row)
            return None

        from core.database.db_manager import SessionLocal
        from core.database.repositories import BlueprintRepository
        db = SessionLocal()
        try:
            repo = BlueprintRepository(db)
            row = repo.get_blueprint(prefix, locale)
            if row:
                return self._resolve_from_db_row(row)
        except Exception as e:
            logger.warning(f"Database resolution failed for stem '{file_stem}': {e}")
        finally:
            db.close()
            
        return None

    def resolve_assets_for_input_file(self, input_file_path: str) -> Optional[InvoiceAssets]:
        """
        Main entry point. Finds the config and template needed to process the given input file.
        
        Args:
            input_file_path: The path to the user's input data file (e.g. 'JF25058.json')
            
        Returns:
            InvoiceAssets object containing resolved configuration and template, or None.
        """
        input_path = Path(input_file_path)
        stem = input_path.stem
        
        logger.info(f"Resolving assets for input: {stem}")

        # Try database resolution
        db_assets = self._try_resolve_from_db(stem)
        if db_assets:
            logger.info(f"✅ Resolved assets using Database Strategy for '{stem}'")
            return InvoiceAssets(
                data_path=input_path,
                config_path=db_assets.config_path,
                template_path=db_assets.template_path,
                config_data=db_assets.config_data,
                template_json_data=db_assets.template_json_data,
                template_xlsx_bytes=db_assets.template_xlsx_bytes
            )

        logger.error(f"❌ Could not resolve any valid assets for {stem}")
        return None

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
        
        match = re.match(r'^([a-zA-Z\-_]+)', stem)
        prefix = match.group(1) if match else None
        if prefix:
            prefix = prefix.rstrip('_')
            
        if not prefix:
            return []
            
        if self._repository:
            db_rows = self._repository.get_customer_variants(prefix)
            variants = []
            for row in db_rows:
                assets = self._resolve_from_db_row(row)
                if assets:
                    variants.append({
                        "suffix": f"_{row.locale}",
                        "config_path": assets.config_path,
                        "template_path": assets.template_path,
                        "config_data": assets.config_data,
                        "template_json_data": assets.template_json_data,
                        "template_xlsx_bytes": assets.template_xlsx_bytes
                    })
            return variants

        from core.database.db_manager import SessionLocal
        from core.database.repositories import BlueprintRepository
        db = SessionLocal()
        try:
            repo = BlueprintRepository(db)
            db_rows = repo.get_customer_variants(prefix)
            if db_rows:
                variants = []
                for row in db_rows:
                    assets = self._resolve_from_db_row(row)
                    if assets:
                        variants.append({
                            "suffix": f"_{row.locale}",
                            "config_path": assets.config_path,
                            "template_path": assets.template_path,
                            "config_data": assets.config_data,
                            "template_json_data": assets.template_json_data,
                            "template_xlsx_bytes": assets.template_xlsx_bytes
                        })
                return variants
        except Exception as e:
            logger.warning(f"Database resolve variants failed: {e}")
        finally:
            db.close()
        
        return []

