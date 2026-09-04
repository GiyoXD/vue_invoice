import logging
import json
import datetime
from fastapi import APIRouter, Depends
from fastapi.responses import JSONResponse, StreamingResponse
from sqlalchemy.orm import Session
from pathlib import Path
from typing import Optional
from core.system_config import sys_config
from core.database.db_manager import get_db
from core.database.repositories.invoice_repository import InvoiceRepository
from core.utils.clock import now as ict_now

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
            except Exception: continue

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
            except Exception: continue

    history.sort(key=lambda x: x["timestamp"] or "", reverse=True)
    return history

@router.get("/history/view")
async def view_history_item(filename: str, source: str = "run_log", db: Session = Depends(get_db)):
    if ".." in filename or "/" in filename or "\\" in filename:
         return JSONResponse(status_code=400, content={"error": "Invalid filename"})

    if source == "processed":
        file_path = sys_config.temp_uploads_dir / "processed" / filename
    elif source == "accepted":
        try:
            repo = InvoiceRepository(db)
            record = repo.get_by_filename(filename)
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
    repo = InvoiceRepository(db)
    existing = repo.get_by_filename(req.filename)
    return {"exists": existing is not None}

@router.post("/registry/accept")
async def accept_invoice(req: HistoryRequest, db: Session = Depends(get_db)):
    if not req.filename or ".." in req.filename or "/" in req.filename or "\\" in req.filename:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})

    processed_dir = sys_config.temp_uploads_dir / "processed"
    file_path = processed_dir / req.filename
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"error": "Pending file not found"})

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as jde:
        return JSONResponse(status_code=422, content={
            "error": f"Cannot accept: data file '{req.filename}' is corrupt or incomplete. Please re-upload and re-process.",
            "details": str(jde)
        })
    except OSError as oe:
        return JSONResponse(status_code=500, content={"error": f"Could not read data file: {str(oe)}"})

    try:
        repo = InvoiceRepository(db)
        repo.accept_processed_invoice(req.filename, data)
        try:
            file_path.unlink()
        except OSError as oe:
            logger.warning(f"Failed to unlink pending file {file_path} after successful accept: {oe}")
        return {"status": "success", "message": f"Invoice {req.filename} accepted."}
    except ValueError as ve:
        return JSONResponse(status_code=422, content={"error": str(ve)})
    except Exception as e:
        logger.error(f"Accept failed: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/registry/reject")
async def reject_invoice(req: HistoryRequest, db: Session = Depends(get_db)):
    if not req.filename or ".." in req.filename or "/" in req.filename or "\\" in req.filename:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})
    file_path = sys_config.temp_uploads_dir / "processed" / req.filename
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"error": "File not found"})
    try:
        repo = InvoiceRepository(db)
        repo.delete_invoice(req.filename)
        file_path.unlink()
        return {"status": "success"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/registry/export")
async def export_registry(start_date: Optional[str] = None, end_date: Optional[str] = None, db: Session = Depends(get_db)):
    try:
        repo = InvoiceRepository(db)
        csv_data = repo.generate_registry_csv(start_date=start_date, end_date=end_date)
        filename = f"export_{ict_now().strftime('%Y%m%d_%H%M%S')}.csv"
        return StreamingResponse(iter([csv_data]), media_type="text/csv", headers={"Content-Disposition": f"attachment; filename={filename}"})
    except ValueError as ve:
        return JSONResponse(status_code=400, content={"error": str(ve)})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/registry/list")
async def list_registry(db: Session = Depends(get_db)):
    try:
        repo = InvoiceRepository(db)
        results = repo.list_recent(limit=10)
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
