import logging
from sqlalchemy import text
from core.database.session import engine, Base

# Import all models to ensure they register on Base.metadata before create_all
from core.database.models.blueprint import Blueprint, BlueprintTemplate
from core.database.models.invoice import ProcessedData, InvoiceItem
from core.database.models.global_map import (
    GlobalMapColumn,
    GlobalMapColumnKeyword,
    GlobalMapHeaderTextMapping,
    GlobalMapSheet,
    GlobalMapFooterLabelKeyword,
    GlobalMapFallbackStrategy,
)
from core.database.models.system import SystemSetting

logger = logging.getLogger(__name__)

def init_db():
    # Pre-check schema for global_map_sheets before creating tables
    try:
        with engine.begin() as conn:
            result = conn.execute(text("PRAGMA table_info(global_map_sheets)")).fetchall()
            columns = [row[1] for row in result]
            if columns and ("canonical_name" in columns or "aliases" in columns or "raw_name" in columns):
                conn.execute(text("DROP TABLE global_map_sheets"))
                print("Dropped legacy global_map_sheets table for schema migration.")
    except Exception as e:
        logger.warning(f"Error checking/dropping legacy global_map_sheets table: {e}")

    try:
        Base.metadata.create_all(bind=engine)
    except Exception as e:
        logger.exception("Failed to initialize database tables (Base.metadata.create_all)")
        raise e
    
    # Auto-migration for newly added columns and legacy tables cleanup
    try:
        with engine.begin() as conn:
            # Drop obsolete legacy tables if they exist
            conn.execute(text("DROP TABLE IF EXISTS global_map_sheet_name_mappings"))
            conn.execute(text("DROP TABLE IF EXISTS global_map_sheet_classifications"))
            
            result = conn.execute(text("PRAGMA table_info(processed_data)")).fetchall()
            columns = [row[1] for row in result]
            if columns and "total_pallets" not in columns:
                conn.execute(text("ALTER TABLE processed_data ADD COLUMN total_pallets FLOAT"))
                print("Successfully added total_pallets column.")
            if columns and "total_net" not in columns:
                conn.execute(text("ALTER TABLE processed_data ADD COLUMN total_net FLOAT"))
                print("Successfully added total_net column.")

            # Dynamic cleanup: Remove legacy sheet classifications from existing database if present
            legacy_sheets_to_remove = ["shipping", "bill", "detail", "content", "weight"]
            for sheet in legacy_sheets_to_remove:
                conn.execute(
                    text("DELETE FROM global_map_sheets WHERE LOWER(TRIM(sheet_name)) = :name"),
                    {"name": sheet}
                )
    except Exception as e:
        logger.exception("Failed to run database migrations and updates")
        raise e
