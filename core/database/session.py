import logging
import os
from pathlib import Path
from sqlalchemy import create_engine, event
from sqlalchemy.orm import declarative_base, sessionmaker
from core.system_config import PROJECT_ROOT

logger = logging.getLogger(__name__)

# Database location: storage/db/invoice_registry.db
DB_DIR = PROJECT_ROOT / "storage" / "db"
DB_DIR.mkdir(parents=True, exist_ok=True)
ENV_DB_PATH = os.getenv("SQLITE_DB_PATH")
if ENV_DB_PATH:
    custom_path = Path(ENV_DB_PATH)
    DB_PATH = custom_path if custom_path.is_absolute() else (PROJECT_ROOT / custom_path).resolve()
else:
    DB_PATH = DB_DIR / "invoice_registry.db"
SQLALCHEMY_DATABASE_URL = f"sqlite:///{DB_PATH.absolute().as_posix()}"
engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})

@event.listens_for(engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA busy_timeout=5000")
    cursor.execute("PRAGMA synchronous=NORMAL")
    cursor.close()

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
