import logging
from fastapi import APIRouter, UploadFile, File, Depends
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from sqlalchemy.orm import Session

from core.database.db_manager import get_db
from core.services.blueprint_service import BlueprintService

router = APIRouter(prefix="/api", tags=["templates"])
logger = logging.getLogger(__name__)

class TemplateConfig(BaseModel):
    customer_code: str
    locale: str = "KH"
    user_mappings: dict
    temp_filename: str
    bundle_dir_name: str = ""
    confirmed_footers: List[str] = []
    pricing_mode: str = "standard"  # 'standard' or 'net'
    ignore_missing_description: bool = False

class CellOverrideRequest(BaseModel):
    customer_code: str
    locale: str = "KH"
    sheet_name: str
    cell_address: str
    overrides: Dict[str, str]

class TemplateNotesRequest(BaseModel):
    customer_code: str
    locale: str = "KH"
    notes: str

class AnalyzeExistingTemplateRequest(BaseModel):
    filename: str
    ignore_missing_description: bool = False

# --- Routes ---

@router.post("/template/analyze")
def analyze_template(
    file: UploadFile = File(...),
    ignore_missing_description: bool = False,
    db: Session = Depends(get_db)
):
    try:
        service = BlueprintService(db)
        res = service.analyze_template_legacy(file.filename, file.file, ignore_missing_description)
        return res
    except Exception as e:
        logger.exception("Analyze template legacy failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/template/analyze-existing")
def analyze_existing_template(req: AnalyzeExistingTemplateRequest, db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        res = service.analyze_template_existing(req.filename, req.ignore_missing_description)
        return res
    except FileNotFoundError as fe:
        return JSONResponse(status_code=404, content={"error": str(fe)})
    except Exception as e:
        logger.exception("Analyze existing template failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/template/generate")
def generate_template(config: TemplateConfig, db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        res = service.generate_blueprint(
            file_token=config.temp_filename,
            customer_code=config.customer_code,
            locale=config.locale,
            mappings=config.user_mappings,
            footer_mappings=config.confirmed_footers,
            pricing_mode=config.pricing_mode,
            ignore_missing_description=config.ignore_missing_description,
            bundle_dir_name=config.bundle_dir_name
        )
        return res
    except Exception as e:
        logger.exception("Template generation legacy failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/templates")
async def list_templates(db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        return service.list_blueprints()
    except Exception as e:
        logger.exception("Listing templates failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/template/view")
async def view_template(customer_code: str, locale: str = "KH", db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        res = service.view_blueprint(customer_code, locale)
        if not res:
            return JSONResponse(status_code=404, content={"error": "Not found"})
        return res
    except Exception as e:
        logger.exception("Viewing template failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.patch("/template/cell")
async def update_template_cell(req: CellOverrideRequest, db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        success = service.update_blueprint_cell(
            customer_code=req.customer_code,
            locale=req.locale,
            sheet_name=req.sheet_name,
            cell_address=req.cell_address,
            overrides=req.overrides
        )
        if not success:
            return JSONResponse(status_code=404, content={"error": "Not found or sheet not found"})
        return {"status": "success"}
    except Exception as e:
        logger.exception("Updating template cell failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.patch("/template/notes")
async def update_template_notes(req: TemplateNotesRequest, db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        success = service.update_blueprint_notes(req.customer_code, req.locale, req.notes)
        if not success:
            return JSONResponse(status_code=404, content={"error": "Not found"})
        return {"status": "success"}
    except Exception as e:
        logger.exception("Updating template notes failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.delete("/template/{customer_code}")
async def delete_template(customer_code: str, locale: str = "KH", db: Session = Depends(get_db)):
    try:
        service = BlueprintService(db)
        success = service.delete_blueprint(customer_code, locale)
        if not success:
            return JSONResponse(status_code=404, content={"error": "Not found"})
        return {"status": "success"}
    except Exception as e:
        logger.exception("Deleting template failed")
        return JSONResponse(status_code=500, content={"error": str(e)})
