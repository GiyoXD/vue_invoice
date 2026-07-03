from sqlalchemy import Column, Integer, String, Float, DateTime, JSON
from core.database.session import Base
from core.utils.clock import now as ict_now

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
    col_pallet_no = Column(String)
    col_net = Column(Float)
    col_gross = Column(Float)
    col_cbm_raw = Column(String)
    col_hs_code = Column(String)
    col_unit_price = Column(Float)
    col_amount = Column(Float)
    is_adjustment = Column(Integer, default=0) # SQLite uses 0/1 for False/True usually, but SQLAlchemy handles Booleans. Let's use Integer for safety or Boolean.
    timestamp = Column(DateTime, default=ict_now)
