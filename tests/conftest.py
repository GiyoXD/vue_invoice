import os
import shutil
import pytest
from pathlib import Path
from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker
from fastapi.testclient import TestClient

# 1. Setup test database engine and session before importing the app
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
        
    # Copy the master database to the test database location to initialize with seeded data
    master_db_path = Path("database/invoice_registry.db")
    if master_db_path.exists():
        shutil.copy2(master_db_path, TEST_DB_PATH)
    else:
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


@pytest.fixture(scope="function", autouse=True)
def db():
    """Provides a database session for testing and resets the database from the master file before execution."""
    # Dispose engine to ensure no locked connections on Windows
    test_engine.dispose()
    
    # Reset test db from master db file before running test
    master_db_path = Path("database/invoice_registry.db")
    if master_db_path.exists():
        shutil.copy2(master_db_path, TEST_DB_PATH)
    else:
        init_db()
        
    session = TestSessionLocal()
    try:
        yield session
    finally:
        session.rollback()
        session.close()
        test_engine.dispose()


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
