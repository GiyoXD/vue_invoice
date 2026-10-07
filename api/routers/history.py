import logging
from typing import Optional
from fastapi import APIRouter, Depends
from fastapi.responses import JSONResponse, StreamingResponse
from sqlalchemy.orm import Session
from pydantic import BaseModel
from core.database.session import get_db
from core.services.shipment import ShipmentService
from core.utils.clock import now as ict_now

router = APIRouter(prefix="/api", tags=["history"])
logger = logging.getLogger(__name__)


class HistoryRequest(BaseModel):
    filename: str


@router.get("/history")
async def get_history(db: Session = Depends(get_db)):
    """Retrieve list of past runs."""
    service = ShipmentService(db)
    return service.get_history()


@router.get("/history/view")
async def view_history_item(filename: str, source: str = "run_log", db: Session = Depends(get_db)):
    if not filename or ".." in filename or "/" in filename or "\\" in filename:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})

    service = ShipmentService(db)
    try:
        return service.get_history_item(filename=filename, source=source)
    except ValueError as ve:
        if ve.__class__.__name__ == "JSONDecodeError":
            return JSONResponse(
                status_code=422,
                content={
                    "error": f"Data file '{filename}' is corrupt or incomplete (truncated JSON). Please re-upload the source Excel file.",
                    "details": str(ve),
                },
            )
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})
    except LookupError:
        return JSONResponse(status_code=404, content={"error": "Database record not found"})
    except FileNotFoundError:
        return JSONResponse(status_code=404, content={"error": "File not found"})
    except OSError as oe:
        return JSONResponse(status_code=500, content={"error": f"Could not read file: {str(oe)}"})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.post("/registry/check")
async def check_invoice_exists(req: HistoryRequest, db: Session = Depends(get_db)):
    if not req.filename or ".." in req.filename or "/" in req.filename or "\\" in req.filename:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})
    try:
        service = ShipmentService(db)
        return {"exists": service.check_exists(req.filename)}
    except ValueError:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.post("/registry/accept")
async def accept_invoice(req: HistoryRequest, db: Session = Depends(get_db)):
    if not req.filename or ".." in req.filename or "/" in req.filename or "\\" in req.filename:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})

    service = ShipmentService(db)
    try:
        service.accept(req.filename)
        return {"status": "success", "message": f"Invoice {req.filename} accepted."}
    except FileNotFoundError:
        return JSONResponse(status_code=404, content={"error": "Pending file not found"})
    except ValueError as ve:
        if ve.__class__.__name__ == "JSONDecodeError":
            return JSONResponse(
                status_code=422,
                content={
                    "error": f"Cannot accept: data file '{req.filename}' is corrupt or incomplete. Please re-upload and re-process.",
                    "details": str(ve),
                },
            )
        return JSONResponse(status_code=422, content={"error": str(ve)})
    except OSError as oe:
        return JSONResponse(status_code=500, content={"error": f"Could not read data file: {str(oe)}"})
    except Exception as e:
        logger.error(f"Accept failed: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.post("/registry/reject")
async def reject_invoice(req: HistoryRequest, db: Session = Depends(get_db)):
    if not req.filename or ".." in req.filename or "/" in req.filename or "\\" in req.filename:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})
    try:
        service = ShipmentService(db)
        service.reject(req.filename)
        return {"status": "success"}
    except FileNotFoundError:
        return JSONResponse(status_code=404, content={"error": "File not found"})
    except ValueError:
        return JSONResponse(status_code=400, content={"error": "Invalid filename"})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.get("/registry/export")
async def export_registry(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    db: Session = Depends(get_db),
):
    try:
        service = ShipmentService(db)
        csv_data = service.export_registry_csv(start_date=start_date, end_date=end_date)
        filename = f"export_{ict_now().strftime('%Y%m%d_%H%M%S')}.csv"
        return StreamingResponse(
            iter([csv_data]),
            media_type="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )
    except ValueError as ve:
        return JSONResponse(status_code=400, content={"error": str(ve)})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.get("/registry/list")
async def list_registry(db: Session = Depends(get_db)):
    try:
        service = ShipmentService(db)
        return service.list_registry(limit=10)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
