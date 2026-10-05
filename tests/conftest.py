import shutil
import pytest
from pathlib import Path
from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker
from fastapi.testclient import TestClient

# 1. Setup test database engine and session before importing the app
import core.database.db_manager as db_manager

TEST_DB_PATH = Path("storage/db/test_invoice_registry.db")
TEST_DB_PATH.parent.mkdir(parents=True, exist_ok=True)

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

# Override the engine and SessionLocal in db_manager and session
db_manager.engine = test_engine
db_manager.SessionLocal = TestSessionLocal
import core.database.session as db_session
db_session.engine = test_engine
db_session.SessionLocal = TestSessionLocal

_db_mutated = False

@event.listens_for(test_engine, "commit")
def _on_test_engine_commit(conn):
    global _db_mutated
    _db_mutated = True

from api.main import app
from core.database.db_manager import init_db, get_db
from core.utils.cache import mapping_cache
from core.services.mapping_service import MappingService

@pytest.fixture(scope="session", autouse=True)
def setup_test_db():
    # Make sure test database file does not exist initially
    if TEST_DB_PATH.exists():
        try:
            TEST_DB_PATH.unlink()
        except Exception:
            pass
        
    # Copy the master database to the test database location to initialize with seeded data
    master_db_path = Path("storage/db/invoice_registry.db")
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
    """Provides a database session for testing and resets the database if modified by test."""
    master_db_path = Path("storage/db/invoice_registry.db")
    if not TEST_DB_PATH.exists():
        if master_db_path.exists():
            shutil.copy2(master_db_path, TEST_DB_PATH)
        else:
            init_db()

    global _db_mutated
    _db_mutated = False
    session = TestSessionLocal()
    try:
        yield session
    finally:
        session.rollback()
        session.close()
        if _db_mutated:
            mapping_cache.invalidate()
            test_engine.dispose()
            if master_db_path.exists():
                shutil.copy2(master_db_path, TEST_DB_PATH)
            else:
                init_db()
            reset_session = TestSessionLocal()
            try:
                MappingService(reset_session).reload_dynamic_state(force=True)
            finally:
                reset_session.close()
            mapping_cache.invalidate()
            _db_mutated = False


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


@pytest.fixture(scope="session")
def test_kh_config() -> dict:
    """Provides the test_KH configuration JSON data from the database or the test_KH_config.json file."""
    config_path = Path(__file__).parent / "fixtures" / "test_KH_config.json"
    if config_path.exists():
        import json
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
            
    import sqlite3
    import json
    conn = sqlite3.connect("storage/db/invoice_registry.db")
    cursor = conn.cursor()
    row = cursor.execute("SELECT config_json FROM blueprints WHERE customer_code = 'test' AND locale = 'KH'").fetchone()
    conn.close()
    if row:
        return json.loads(row[0])
    return {}


@pytest.fixture(scope="session")
def test_kh_template() -> dict:
    """Provides the test_KH template JSON data from the database or the test_KH_template.json file."""
    template_path = Path(__file__).parent / "fixtures" / "test_KH_template.json"
    if template_path.exists():
        import json
        with open(template_path, "r", encoding="utf-8") as f:
            return json.load(f)
            
    import sqlite3
    import json
    conn = sqlite3.connect("storage/db/invoice_registry.db")
    cursor = conn.cursor()
    row = cursor.execute("SELECT template_json FROM blueprints WHERE customer_code = 'test' AND locale = 'KH'").fetchone()
    conn.close()
    if row:
        return json.loads(row[0])
    return {}

