import logging
import json
import datetime
import io
import csv
from fastapi import APIRouter, Depends
from fastapi.responses import JSONResponse, StreamingResponse
from sqlalchemy.orm import Session
from pathlib import Path
from typing import Optional, List
from core.system_config import sys_config
from core.database.db_manager import get_db, ProcessedData, InvoiceItem, get_cambodia_time, init_db, engine, Base

router = APIRouter(prefix="/api", tags=["history"])
logger = logging.getLogger(__name__)

from pydantic import BaseModel
class HistoryRequest(BaseModel):
    filename: str

@router.get("/history")
async def get_history():
    """Retrieve list of past runs."""
    history = []
    
    run_log_dir = Path("run_log")
    if run_log_dir.exists():
        for f in run_log_dir.glob("*_metadata.json"):
            try:
                with open(f, 'r', encoding='utf-8') as json_file:
                    data = json.load(json_file)
                    history.append({
                        "filename": f.name,
                        "type": "run_log",
                        "timestamp": data.get("timestamp"),
                        "output_file": data.get("output_file"),
                        "status": data.get("status"),
                        "item_count": data.get("database_export", {}).get("summary", {}).get("item_count", 0),
                        "total_sqft": data.get("database_export", {}).get("summary", {}).get("total_sqft", 0)
                    })
            except: continue

    processed_dir = sys_config.temp_uploads_dir / "processed"
    if processed_dir.exists():
        for f in processed_dir.glob("*.json"):
            try:
                if any(h["filename"] == f.name for h in history): continue
                stats = f.stat()
                with open(f, 'r', encoding='utf-8') as json_file:
                    data = json.load(json_file)
                    item_count = 0
                    if "multi_table" in data and data["multi_table"]:
                        item_count = sum(len(table) for table in data["multi_table"] if isinstance(table, list))
                    
                    history.append({
                        "filename": f.name,
                        "type": "processed",
                        "timestamp": datetime.datetime.fromtimestamp(stats.st_mtime).isoformat(),
                        "output_file": f.name,
                        "status": "Ready",
                        "item_count": item_count,
                        "total_sqft": 0
                    })
            except: continue

    history.sort(key=lambda x: x["timestamp"] or "", reverse=True)
    return history

@router.get("/history/view")
async def view_history_item(filename: str, source: str = "run_log"):
    if ".." in filename or "/" in filename or "\\" in filename:
         return JSONResponse(status_code=400, content={"error": "Invalid filename"})

    if source == "processed":
        file_path = sys_config.temp_uploads_dir / "processed" / filename
    elif source == "accepted":
        try:
            db = next(get_db())
            record = db.query(ProcessedData).filter(ProcessedData.filename == filename).first()
            if record: return record.data_payload
            return JSONResponse(status_code=404, content={"error": "Database record not found"})
        except Exception as e:
            return JSONResponse(status_code=500, content={"error": str(e)})
    else:
        file_path = Path("run_log") / filename
        if not file_path.exists():
            fallback_path = sys_config.temp_uploads_dir / "processed" / filename
            if fallback_path.exists(): file_path = fallback_path
    
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"error": "File not found"})
        
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as jde:
        return JSONResponse(status_code=422, content={
            "error": f"Data file '{filename}' is corrupt or incomplete (truncated JSON). Please re-upload the source Excel file.",
            "details": str(jde)
        })
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Could not read file: {str(e)}"})

@router.post("/registry/check")
async def check_invoice_exists(req: HistoryRequest, db: Session = Depends(get_db)):
    existing = db.query(ProcessedData).filter(ProcessedData.filename == req.filename).first()
    return {"exists": existing is not None}

@router.post("/registry/accept")
async def accept_invoice(req: HistoryRequest, db: Session = Depends(get_db)):
    processed_dir = sys_config.temp_uploads_dir / "processed"
    file_path = processed_dir / req.filename
    if not file_path.exists(): return JSONResponse(status_code=404, content={"error": "Pending file not found"})
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as jde:
        return JSONResponse(status_code=422, content={
            "error": f"Cannot accept: data file '{req.filename}' is corrupt or incomplete. Please re-upload and re-process.",
            "details": str(jde)
        })
    try:
        db.query(InvoiceItem).filter(InvoiceItem.invoice_id == req.filename).delete()
        
        items_to_add = []
        item_count = 0
        total_sqft = 0.0
        total_net = 0.0
        total_amount = 0.0
        total_pallets = 0.0
        
        def to_float(val):
            try:
                if isinstance(val, str): val = val.replace(',', '').replace('$', '').strip()
                return float(val) if val is not None and str(val).strip() != "" else 0.0
            except: return 0.0

        footer_data = data.get("footer_data")
        if not footer_data or "grand_total" not in footer_data:
            return JSONResponse(status_code=422, content={"error": f"Cannot accept: missing footer_data or grand_total in '{req.filename}'"})

        grand_total = footer_data["grand_total"]
        total_pallets = to_float(grand_total.get("col_pallet_count", 0))
        total_sqft = to_float(grand_total.get("col_qty_sf", 0))
        total_net = to_float(grand_total.get("col_net", 0))
        total_amount = to_float(grand_total.get("col_amount", 0))
        
        # Prioritize multi_table so parsed values (like pallets and CBM) are accurately tracked
        source_tables = data.get("multi_table") or data.get("raw_data") or []
        for table in source_tables:
            if isinstance(table, list):
                for row in table:
                    sqft_val = to_float(row.get("col_qty_sf"))
                    amount_val = to_float(row.get("col_amount"))
                    item_count += 1

                    item = InvoiceItem(
                        invoice_id=req.filename,
                        col_dc=str(row.get("col_dc", "")),
                        col_po=str(row.get("col_po", "")),
                        col_production_order_no=str(row.get("col_production_order_no", "")),
                        col_production_date=str(row.get("col_production_date", "")),
                        col_line_no=str(row.get("col_line_no", "")),
                        col_direction=str(row.get("col_direction", "")),
                        col_item=str(row.get("col_item", "")),
                        col_reference_code=str(row.get("col_reference_code", "")),
                        col_desc=str(row.get("col_desc", "")),
                        col_level=str(row.get("col_level", "")),
                        col_grade=str(row.get("col_grade", "")),
                        col_qty_pcs=to_float(row.get("col_qty_pcs")),
                        col_qty_sf=sqft_val,
                        col_pallet_count=to_float(row.get("col_pallet_count")),
                        col_net=to_float(row.get("col_net")),
                        col_gross=to_float(row.get("col_gross")),
                        col_cbm_raw=str(row.get("col_cbm_raw", row.get("col_cbm", ""))),
                        col_hs_code=str(row.get("col_hs_code", "")),
                        col_unit_price=to_float(row.get("col_unit_price")),
                        col_amount=amount_val,
                        is_adjustment=0,
                        timestamp=get_cambodia_time()
                    )
                    items_to_add.append(item)

        if "price_adjustment" in data:
            for adj in data["price_adjustment"]:
                def to_float(val):
                    try: return float(val) if val is not None and str(val).strip() != "" else 0.0
                    except: return 0.0
                amt = to_float(adj.get("amount", 0.0))
                total_amount += amt
                item = InvoiceItem(
                    invoice_id=req.filename,
                    col_desc=str(adj.get("description", "")),
                    col_amount=amt,
                    is_adjustment=1,
                    timestamp=get_cambodia_time()
                )
                items_to_add.append(item)
        
        if items_to_add: db.add_all(items_to_add)
        existing = db.query(ProcessedData).filter(ProcessedData.filename == req.filename).first()
        if existing:
            existing.item_count, existing.total_sqft, existing.total_net, existing.total_amount, existing.total_pallets = item_count, total_sqft, total_net, total_amount, total_pallets
            existing.data_payload, existing.timestamp = data, get_cambodia_time()
        else:
            db.add(ProcessedData(filename=req.filename, item_count=item_count, total_sqft=total_sqft, total_net=total_net, total_amount=total_amount, total_pallets=total_pallets, data_payload=data))
        db.commit()
        file_path.unlink()
        return {"status": "success", "message": f"Invoice {req.filename} accepted."}
    except Exception as e:
        logger.error(f"Accept failed: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/registry/reject")
async def reject_invoice(req: HistoryRequest):
    file_path = sys_config.temp_uploads_dir / "processed" / req.filename
    if not file_path.exists(): return JSONResponse(status_code=404, content={"error": "File not found"})
    try:
        file_path.unlink()
        return {"status": "success"}
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/registry/export")
async def export_registry(start_date: Optional[str] = None, end_date: Optional[str] = None, db: Session = Depends(get_db)):
    try:
        query = db.query(InvoiceItem)
        if start_date: query = query.filter(InvoiceItem.timestamp >= datetime.datetime.fromisoformat(start_date))
        if end_date:
            end_dt = datetime.datetime.fromisoformat(end_date)
            if len(end_date) <= 10: end_dt += datetime.timedelta(days=1)
            query = query.filter(InvoiceItem.timestamp < end_dt)
        results = query.order_by(InvoiceItem.timestamp.desc(), InvoiceItem.invoice_id).all()
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow([
            "ID", "Invoice ID", "Timestamp", "DC", "PO", "Production Order No", 
            "Production Date", "Line No", "Direction", "Item", "Reference Code", 
            "Description", "Level", "Grade", "Qty (PCS)", "Qty (SF)", "Pallet Count", 
            "Net Weight", "Gross Weight", "CBM Raw", 
            "HS Code", "Unit Price", "Amount", "Is Adjustment"
        ])
        
        total_sqft = 0.0
        total_net = 0.0
        total_pallets = 0.0
        total_amount = 0.0
        
        for row in results:
            writer.writerow([
                row.id,
                row.invoice_id.replace(".json", ""),
                row.timestamp.isoformat(),
                row.col_dc,
                row.col_po,
                row.col_production_order_no,
                row.col_production_date,
                row.col_line_no,
                row.col_direction,
                row.col_item,
                row.col_reference_code,
                row.col_desc,
                row.col_level,
                row.col_grade,
                row.col_qty_pcs,
                row.col_qty_sf,
                row.col_pallet_count,
                row.col_net,
                row.col_gross,
                row.col_cbm_raw,
                row.col_hs_code,
                row.col_unit_price,
                row.col_amount,
                "Yes" if row.is_adjustment else "No"
            ])
            
            total_sqft += float(row.col_qty_sf or 0.0)
            total_net += float(row.col_net or 0.0)
            total_pallets += float(row.col_pallet_count or 0.0)
            total_amount += float(row.col_amount or 0.0)
            
        # Write Summary Row
        summary_row = [""] * 24
        summary_row[1] = "TOTAL"
        summary_row[15] = round(total_sqft, 2)
        summary_row[16] = round(total_pallets, 2)
        summary_row[17] = round(total_net, 2)
        summary_row[22] = round(total_amount, 2)
        
        writer.writerow([])
        writer.writerow(summary_row)
        
        output.seek(0)
        csv_data = "\ufeff" + output.getvalue()
        filename = f"export_{get_cambodia_time().strftime('%Y%m%d_%H%M%S')}.csv"
        return StreamingResponse(iter([csv_data]), media_type="text/csv", headers={"Content-Disposition": f"attachment; filename={filename}"})
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/registry/list")
async def list_registry(db: Session = Depends(get_db)):
    try:
        results = db.query(ProcessedData).order_by(ProcessedData.timestamp.desc()).limit(10).all()
        return [
            {
                "id": r.id, 
                "filename": r.filename, 
                "timestamp": r.timestamp.isoformat() if r.timestamp else None, 
                "item_count": r.item_count, 
                "total_sqft": r.total_sqft,
                "total_net": getattr(r, "total_net", 0.0),
                "total_pallets": getattr(r, "total_pallets", 0.0),
                "total_amount": r.total_amount
            } for r in results
        ]
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/registry/reset")
async def reset_registry(db: Session = Depends(get_db)):
    try:
        Base.metadata.drop_all(bind=engine)
        init_db()
        processed_dir = sys_config.temp_uploads_dir / "processed"
        if processed_dir.exists():
            for f in processed_dir.glob("*.json"): f.unlink()
        return {"status": "success"}
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})
