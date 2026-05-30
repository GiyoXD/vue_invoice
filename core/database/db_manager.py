import datetime
import json
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, Text, JSON, ForeignKey, LargeBinary, event
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from pathlib import Path
from core.utils.clock import now as ict_now


# Database location: database/invoice_registry.db
DB_DIR = Path("database")
DB_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DB_DIR / "invoice_registry.db"

# SQLAlchemy Setup
SQLALCHEMY_DATABASE_URL = f"sqlite:///{DB_PATH.absolute()}"
engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})

@event.listens_for(engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

class Blueprint(Base):
    __tablename__ = "blueprints"

    id = Column(Integer, primary_key=True, index=True)
    customer_code = Column(String, nullable=False)
    locale = Column(String, nullable=False, default="KH")
    description = Column(Text, nullable=True)
    config_json = Column(Text, nullable=False)
    template_json = Column(Text, nullable=False)
    created_at = Column(DateTime, default=ict_now)
    updated_at = Column(DateTime, default=ict_now, onupdate=ict_now)

    template_binary = relationship("BlueprintTemplate", uselist=False, back_populates="blueprint", cascade="all, delete-orphan")

class BlueprintTemplate(Base):
    __tablename__ = "blueprint_templates"

    blueprint_id = Column(Integer, ForeignKey("blueprints.id", ondelete="CASCADE"), primary_key=True)
    filename = Column(String, nullable=False)
    xlsx_blob = Column(LargeBinary, nullable=False)

    blueprint = relationship("Blueprint", back_populates="template_binary")

class ProcessedData(Base):
    __tablename__ = "processed_data"

    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String, unique=True, index=True)
    timestamp = Column(DateTime, default=ict_now)
    item_count = Column(Integer)
    total_sqft = Column(Float)
    total_net = Column(Float)
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
    col_net = Column(Float)
    col_gross = Column(Float)
    col_cbm_raw = Column(String)
    col_hs_code = Column(String)
    col_unit_price = Column(Float)
    col_amount = Column(Float)
    is_adjustment = Column(Integer, default=0) # SQLite uses 0/1 for False/True usually, but SQLAlchemy handles Booleans. Let's use Integer for safety or Boolean.
    timestamp = Column(DateTime, default=ict_now)

class SystemSetting(Base):
    __tablename__ = "system_settings"

    key = Column(String, primary_key=True)
    value_json = Column(Text, nullable=False)


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
            if columns and "total_net" not in columns:
                conn.execute(text("ALTER TABLE processed_data ADD COLUMN total_net FLOAT"))
                print("Successfully added total_net column.")
    except Exception as e:
        pass

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def get_global_mapping_config(db=None) -> dict:
    from core.system_config import sys_config
    
    # Load fallback from local disk mapping_config.json first if needed
    disk_data = {}
    try:
        disk_path = sys_config.mapping_config_path
        if disk_path.exists():
            with open(disk_path, 'r', encoding='utf-8') as f:
                disk_data = json.load(f)
    except Exception:
        pass

    close_session = False
    if db is None:
        db = SessionLocal()
        close_session = True
        
    try:
        row = db.query(SystemSetting).filter(SystemSetting.key == "mapping_config").first()
        if row:
            return json.loads(row.value_json)
        else:
            # Seed DB from disk fallback or empty dict
            if disk_data:
                new_row = SystemSetting(key="mapping_config", value_json=json.dumps(disk_data, ensure_ascii=False))
                db.add(new_row)
                db.commit()
                return disk_data
    except Exception as e:
        # If DB query fails (e.g. table not created yet), return disk_data or empty dict
        pass
    finally:
        if close_session:
            db.close()
            
    return disk_data

def save_global_mapping_config(data: dict, db=None) -> None:
    from core.system_config import sys_config
    
    # Save to disk backup
    try:
        disk_path = sys_config.mapping_config_path
        disk_path.parent.mkdir(parents=True, exist_ok=True)
        with open(disk_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

    close_session = False
    if db is None:
        db = SessionLocal()
        close_session = True
        
    try:
        row = db.query(SystemSetting).filter(SystemSetting.key == "mapping_config").first()
        json_str = json.dumps(data, ensure_ascii=False)
        if row:
            row.value_json = json_str
        else:
            row = SystemSetting(key="mapping_config", value_json=json_str)
            db.add(row)
        db.commit()
    except Exception as e:
        db.rollback()
        raise e
    finally:
        if close_session:
            db.close()
