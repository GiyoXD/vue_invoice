import os
import shutil
import tempfile
import pytest
from pathlib import Path
from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker
from fastapi.testclient import TestClient

# 1. Isolate the mapping_config.json file to prevent tests from modifying the git-tracked file
ORIGINAL_MAPPING_CONFIG = Path("database/blueprints/mapper/mapping_config.json").resolve()
temp_dir = tempfile.mkdtemp()
TEMP_MAPPING_CONFIG = Path(temp_dir) / "mapping_config.json"
shutil.copy2(ORIGINAL_MAPPING_CONFIG, TEMP_MAPPING_CONFIG)
os.environ["MAPPING_CONFIG"] = str(TEMP_MAPPING_CONFIG)

# 2. Setup test database engine and session before importing the app
import core.database.db_manager as db_manager

TEST_DB_PATH = Path("database/test_invoice_registry.db")

# Force using the test database engine and session local
test_engine = create_engine(
    f"sqlite:///{TEST_DB_PATH.absolute()}",
    connect_args={"check_same_thread": False}
)
TestSessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=test_engine)

@event.listens_for(test_engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

# Override the engine and SessionLocal in db_manager
db_manager.engine = test_engine
db_manager.SessionLocal = TestSessionLocal

from api.main import app
from core.database.db_manager import init_db, get_db

@pytest.fixture(scope="session", autouse=True)
def setup_test_db():
    # Make sure test database file does not exist initially
    if TEST_DB_PATH.exists():
        try:
            TEST_DB_PATH.unlink()
        except Exception:
            pass
        
    # Initialize the test database schema
    init_db()
    
    yield
    
    # Dispose of engine to release file handles on Windows
    test_engine.dispose()
    
    # Teardown: Clean up the test database file
    if TEST_DB_PATH.exists():
        try:
            TEST_DB_PATH.unlink()
        except Exception:
            pass

    # Clean up the temporary mapping config file and directory
    try:
        if TEMP_MAPPING_CONFIG.exists():
            TEMP_MAPPING_CONFIG.unlink()
        Path(temp_dir).rmdir()
    except Exception:
        pass

@pytest.fixture(scope="function", autouse=True)
def reset_mapping_config_on_disk():
    yield
    # Restore the temporary mapping config file to original defaults after each test
    try:
        shutil.copy2(ORIGINAL_MAPPING_CONFIG, TEMP_MAPPING_CONFIG)
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning(f"Failed to reset mapping config: {e}")


@pytest.fixture(scope="function")
def db():
    """Provides a database session for testing and cleans up test data tables after execution."""
    session = TestSessionLocal()
    try:
        yield session
    finally:
        session.rollback()
        from core.database.db_manager import (
            Blueprint, BlueprintTemplate, ProcessedData, InvoiceItem,
            GlobalMapSheet,
            GlobalMapHeaderTextMapping, GlobalMapColumnKeyword,
            GlobalMapColumn, GlobalMapFooterLabelKeyword, GlobalMapFallbackStrategy
        )
        try:
            session.query(BlueprintTemplate).delete()
            session.query(Blueprint).delete()
            session.query(ProcessedData).delete()
            session.query(InvoiceItem).delete()
            # Clean up global mapping tables (children first)
            session.query(GlobalMapSheet).delete()
            session.query(GlobalMapHeaderTextMapping).delete()
            session.query(GlobalMapColumnKeyword).delete()
            session.query(GlobalMapColumn).delete()
            session.query(GlobalMapFooterLabelKeyword).delete()
            session.query(GlobalMapFallbackStrategy).delete()
            session.commit()
        except Exception:
            session.rollback()
        finally:
            session.close()


@pytest.fixture(scope="function")
def client(db):
    """Provides a TestClient for integration testing, overriding get_db dependency."""
    def override_get_db():
        try:
            yield db
        finally:
            pass
            
    app.dependency_overrides[get_db] = override_get_db
    with TestClient(app) as test_client:
        yield test_client
    app.dependency_overrides.clear()
