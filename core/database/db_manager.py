import datetime
import json
import logging
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, Text, JSON, ForeignKey, LargeBinary, event, UniqueConstraint
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from pathlib import Path
from core.utils.clock import now as ict_now

logger = logging.getLogger(__name__)


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


# --- Global Map Relational Models ---

class GlobalMapColumn(Base):
    __tablename__ = "global_map_columns"
    
    col_id = Column(String(64), primary_key=True)
    excel_format = Column(String(64), default="@")
    
    keywords = relationship("GlobalMapColumnKeyword", back_populates="column", cascade="all, delete-orphan")
    overrides = relationship("GlobalMapHeaderTextMapping", back_populates="column", cascade="all, delete-orphan")

class GlobalMapColumnKeyword(Base):
    __tablename__ = "global_map_column_keywords"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    col_id = Column(String(64), ForeignKey("global_map_columns.col_id", ondelete="CASCADE"), nullable=False)
    keyword = Column(String(128), nullable=False)
    
    column = relationship("GlobalMapColumn", back_populates="keywords")
    
    __table_args__ = (
        UniqueConstraint("col_id", "keyword", name="uq_global_map_col_keyword"),
    )

class GlobalMapHeaderTextMapping(Base):
    __tablename__ = "global_map_header_text_mappings"
    
    raw_text = Column(String(256), primary_key=True)
    canonical_col_id = Column(
        String(64), 
        ForeignKey("global_map_columns.col_id", ondelete="CASCADE"), 
        nullable=False
    )
    
    column = relationship("GlobalMapColumn", back_populates="overrides")

class GlobalMapSheet(Base):
    __tablename__ = "global_map_sheets"
    
    sheet_name = Column(String(128), primary_key=True)
    processing_type = Column(String(64), nullable=False)  # 'aggregation' or 'processed_tables'

class GlobalMapFooterLabelKeyword(Base):
    __tablename__ = "global_map_footer_label_keywords"
    
    keyword = Column(String(128), primary_key=True)

class GlobalMapFallbackStrategy(Base):
    __tablename__ = "global_map_fallback_strategies"
    
    strategy_key = Column(String(128), primary_key=True)
    strategy_value = Column(String(256), nullable=False)



from sqlalchemy import text

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

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def get_global_mapping_config(db=None) -> dict:
    close_session = False
    if db is None:
        db = SessionLocal()
        close_session = True
        
    try:
        # Load from relational tables and reconstruct dictionary
        config = {}
        
        # 1 & 2. Load sheets and classifications from unified table
        sheets = db.query(GlobalMapSheet).all()
        agg_set = set()
        proc_set = set()
        config["sheet_name_mappings"] = {"mappings": {}}
        
        for s in sheets:
            sheet_lower = s.sheet_name.lower().strip()
            if s.processing_type == "aggregation":
                agg_set.add(sheet_lower)
            elif s.processing_type == "processed_tables":
                proc_set.add(sheet_lower)
                        
        config["aggregation_sheets"] = sorted(list(agg_set))
        config["processed_tables_sheets"] = sorted(list(proc_set))
        
        # 3. Load header text mappings
        header_maps = db.query(GlobalMapHeaderTextMapping).all()
        config["header_text_mappings"] = {
            "mappings": {hm.raw_text: hm.canonical_col_id for hm in header_maps}
        }
        
        # 4. Load shipping header map (columns + keywords)
        columns = db.query(GlobalMapColumn).all()
        config["shipping_header_map"] = {}
        for col in columns:
            config["shipping_header_map"][col.col_id] = {
                "keywords": [kw.keyword for kw in col.keywords],
                "format": col.excel_format
            }
            
        # 5. Load footer label mappings
        footers = db.query(GlobalMapFooterLabelKeyword).all()
        config["footer_label_mappings"] = {
            "keywords": [f.keyword for f in footers]
        }
        
        # 6. Load fallback strategies
        strategies = db.query(GlobalMapFallbackStrategy).all()
        config["fallback_strategies"] = {}
        for s in strategies:
            val = s.strategy_value
            if val.lower() == "true": val = True
            elif val.lower() == "false": val = False
            elif val.replace('.', '', 1).isdigit(): val = float(val) if '.' in val else int(val)
            config["fallback_strategies"][s.strategy_key] = val
            
        return config

    except Exception as e:
        logger.exception("Database query failed in get_global_mapping_config")
        raise e
    finally:
        if close_session:
            db.close()

def save_global_mapping_config(data: dict, db=None) -> None:
    close_session = False
    if db is None:
        db = SessionLocal()
        close_session = True
        
    try:
        # 1. Clear existing relational mapping data to overwrite completely
        # Delete children first to satisfy foreign key constraints
        db.query(GlobalMapSheet).delete()
        db.query(GlobalMapHeaderTextMapping).delete()
        db.query(GlobalMapColumnKeyword).delete()
        db.query(GlobalMapColumn).delete()
        db.query(GlobalMapFooterLabelKeyword).delete()
        db.query(GlobalMapFallbackStrategy).delete()
        
        # 2 & 3. Insert sheet classifications into unified table
        agg_sheets = data.get("aggregation_sheets", [])
        proc_sheets = data.get("processed_tables_sheets", [])
        
        seen_sheets = set()
        
        for sheet in agg_sheets:
            sheet_clean = sheet.strip()
            if not sheet_clean:
                continue
            sheet_lower = sheet_clean.lower()
            if sheet_lower not in seen_sheets:
                seen_sheets.add(sheet_lower)
                db.add(GlobalMapSheet(
                    sheet_name=sheet_clean,
                    processing_type="aggregation"
                ))
                
        for sheet in proc_sheets:
            sheet_clean = sheet.strip()
            if not sheet_clean:
                continue
            sheet_lower = sheet_clean.lower()
            if sheet_lower not in seen_sheets:
                seen_sheets.add(sheet_lower)
                db.add(GlobalMapSheet(
                    sheet_name=sheet_clean,
                    processing_type="processed_tables"
                ))
            
        # 4. Insert shipping columns and column keywords
        shipping_map = data.get("shipping_header_map", {})
        for col_id, props in shipping_map.items():
            if not isinstance(props, dict):
                continue
            fmt = props.get("format", "@")
            col = GlobalMapColumn(col_id=col_id, excel_format=fmt)
            db.add(col)
            
            # Keywords
            kws = props.get("keywords", [])
            seen_kws = set()
            for kw in kws:
                kw_clean = kw.strip()
                if not kw_clean:
                    continue
                kw_lower = kw_clean.lower()
                if kw_lower not in seen_kws:
                    seen_kws.add(kw_lower)
                    db.add(GlobalMapColumnKeyword(col_id=col_id, keyword=kw_clean))
                
        # 5. Insert header text mappings (overrides)
        # We only write overrides pointing to column IDs that exist in the columns list
        header_mappings = data.get("header_text_mappings", {}).get("mappings", {})
        for raw, canonical in header_mappings.items():
            if canonical in shipping_map:
                db.add(GlobalMapHeaderTextMapping(raw_text=raw, canonical_col_id=canonical))
                
        # 6. Insert footer label mappings
        footer_kws = data.get("footer_label_mappings", {}).get("keywords", [])
        seen_footers = set()
        for kw in footer_kws:
            kw_clean = kw.strip()
            if not kw_clean:
                continue
            kw_lower = kw_clean.lower()
            if kw_lower not in seen_footers:
                seen_footers.add(kw_lower)
                db.add(GlobalMapFooterLabelKeyword(keyword=kw_clean))
            
        # 7. Insert fallback strategies
        strategies = data.get("fallback_strategies", {})
        for key, val in strategies.items():
            db.add(GlobalMapFallbackStrategy(strategy_key=key, strategy_value=str(val)))
            
        db.commit()
    except Exception as e:
        db.rollback()
        raise e
    finally:
        if close_session:
            db.close()

