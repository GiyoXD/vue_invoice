import datetime
import json
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, Text, JSON
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from pathlib import Path

def get_cambodia_time():
    return datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7)))


# Database location: database/invoice_registry.db
DB_DIR = Path("database")
DB_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DB_DIR / "invoice_registry.db"

# SQLAlchemy Setup
SQLALCHEMY_DATABASE_URL = f"sqlite:///{DB_PATH.absolute()}"
engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

class ProcessedData(Base):
    __tablename__ = "processed_data"

    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String, unique=True, index=True)
    timestamp = Column(DateTime, default=get_cambodia_time)
    item_count = Column(Integer)
    total_sqft = Column(Float)
    total_amount = Column(Float)
    total_pallets = Column(Float)
    # Storing the full JSON payload
    data_payload = Column(JSON)
    status = Column(String, default="Accepted")

class InvoiceItem(Base):
    __tablename__ = "invoice_items"

    id = Column(Integer, primary_key=True, index=True)
    invoice_id = Column(String, index=True) # Maps to ProcessedData.filename
    col_dc = Column(String)
    col_po = Column(String)
    col_production_order_no = Column(String)
    col_production_date = Column(String)
    col_line_no = Column(String)
    col_direction = Column(String)
    col_item = Column(String)
    col_reference_code = Column(String)
    col_desc = Column(String)
    col_level = Column(String)
    col_grade = Column(String)
    col_qty_pcs = Column(Float)
    col_qty_sf = Column(Float)
    col_pallet_count = Column(Float)
    col_pallet_count_raw = Column(String)
    col_net = Column(Float)
    col_gross = Column(Float)
    col_cbm_raw = Column(String)
    col_hs_code = Column(String)
    col_unit_price = Column(Float)
    col_amount = Column(Float)
    is_adjustment = Column(Integer, default=0) # SQLite uses 0/1 for False/True usually, but SQLAlchemy handles Booleans. Let's use Integer for safety or Boolean.
    timestamp = Column(DateTime, default=get_cambodia_time)


from sqlalchemy import text

def init_db():
    Base.metadata.create_all(bind=engine)
    
    # Auto-migration for newly added columns
    try:
        with engine.begin() as conn:
            result = conn.execute(text("PRAGMA table_info(processed_data)")).fetchall()
            columns = [row[1] for row in result]
            if columns and "total_pallets" not in columns:
                conn.execute(text("ALTER TABLE processed_data ADD COLUMN total_pallets FLOAT"))
                print("Successfully added total_pallets column.")
    except Exception as e:
        pass

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
