from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from sqlalchemy.orm import Session
from typing import Dict, List, Optional, Any
from pathlib import Path
import shutil
import json
import logging

from core.database.db_manager import get_db
from core.services.blueprint_service import BlueprintService
from core.services.mapping_service import MappingService

router = APIRouter(prefix="/api/blueprint", tags=["blueprint"])
logger = logging.getLogger(__name__)

# --- Schemas ---

class ScanResult(BaseModel):
    status: str  # "clean" or "needs_mapping"
    file_token: str # Temporary filename to reference in step 2
    unknown_headers: List[str] = []
    unconfirmed_footers: List[str] = []
    unrecognized_sheets: List[str] = []
    warnings: List[str] = []
    preview_analysis: Optional[Dict[str, Any]] = None

class GenerateRequest(BaseModel):
    file_token: str
    customer_code: str # e.g. "CLW"
    locale: str = "KH"
    mappings: Dict[str, str] = {} # {"Unknown Header": "col_remark"}
    footer_mappings: List[str] = []
    pricing_mode: str = "standard"  # 'standard' or 'net'

class GenerateResult(BaseModel):
    status: str
    message: str

# --- Endpoints ---

@router.post("/scan", response_model=ScanResult)
async def scan_template(file: UploadFile = File(...), db: Session = Depends(get_db)):
    """
    Step 1: Scan uploaded template.
    Returns 'needs_mapping' if unknown columns are found.
    """
    try:
        service = BlueprintService(db)
        res = service.scan_and_analyze_template(file.filename, file.file)
        return res
    except Exception as e:
        logger.error(f"Scan failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate", response_model=GenerateResult)
async def generate_config(request: GenerateRequest, db: Session = Depends(get_db)):
    """
    Step 2: Generate final config using verified/mapped headers.
    """
    try:
        service = BlueprintService(db)
        res = service.generate_blueprint(
            file_token=request.file_token,
            customer_code=request.customer_code,
            locale=request.locale,
            mappings=request.mappings,
            footer_mappings=request.footer_mappings,
            pricing_mode=request.pricing_mode
        )
        return res
    except Exception as e:
        logger.error(f"Blueprint generation failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/options")
async def get_mapping_options():
    """
    Return list of valid system columns for mapping.
    Frontend uses this to populate the dropdown.
    """
    return BlueprintService.get_mapping_options()

@router.get("/mappings")
async def get_mappings(mapping_type: str = "header_text_mappings", db: Session = Depends(get_db)):
    """
    Get the global mapping dictionary of the specified type.
    Options: header_text_mappings, sheet_name_mappings, shipping_header_map
    """
    try:
        service = MappingService(db)
        return service.get_mappings(mapping_type)
    except Exception as e:
        logger.error(f"Failed to get mappings: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

class MappingsUpdateRequest(BaseModel):
    mapping_type: str = "header_text_mappings"
    mappings: Dict[str, Any]

@router.post("/mappings")
async def update_mappings(request: MappingsUpdateRequest, db: Session = Depends(get_db)):
    """
    Overwrite the specified global mapping dictionary.
    """
    try:
        service = MappingService(db)
        service.update_mappings(request.mapping_type, request.mappings)
        return {"status": "success"}
    except Exception as e:
        logger.error(f"Failed to update mappings: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})
