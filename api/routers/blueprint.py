from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import Dict, List, Optional, Any
from pathlib import Path
import shutil
import json
import logging

from core.system_config import sys_config
from core.orchestrator import Orchestrator

router = APIRouter(prefix="/api/blueprint", tags=["blueprint"])
logger = logging.getLogger(__name__)

# --- Schemas ---

class ScanResult(BaseModel):
    status: str  # "clean" or "needs_mapping"
    file_token: str # Temporary filename to reference in step 2
    unknown_headers: List[str] = []
    preview_analysis: Optional[Dict[str, Any]] = None

class GenerateRequest(BaseModel):
    file_token: str
    customer_code: str # e.g. "CLW"
    mappings: Dict[str, str] = {} # {"Unknown Header": "col_remark"}

class GenerateResult(BaseModel):
    status: str
    config_path: str
    template_path: str
    message: str

# --- Endpoints ---

@router.post("/scan", response_model=ScanResult)
async def scan_template(file: UploadFile = File(...)):
    """
    Step 1: Scan uploaded template.
    Returns 'needs_mapping' if unknown columns are found.
    """
    temp_dir = sys_config.temp_uploads_dir
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    file_token = f"scan_{file.filename}"
    file_path = temp_dir / file_token
    
    try:
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Run Scanner (We'll use Orchestrator logic here)
        orchestrator = Orchestrator()
        analysis_json_str = orchestrator.analyze_template(file_path)
        analysis = json.loads(analysis_json_str)
        
        # Check for unknowns
        unknown_headers = []
        for sheet in analysis.get("sheets", []):
            for header in sheet.get("header_positions", []):
                if StringUtils.is_unknown_col_id(header.get("col_id")):
                    unknown_headers.append(header.get("keyword"))
        
        # Deduplicate
        unknown_headers = list(set(unknown_headers))
        
        if unknown_headers:
            return ScanResult(
                status="needs_mapping",
                file_token=file_token,
                unknown_headers=unknown_headers,
                preview_analysis=analysis
            )
        else:
             return ScanResult(
                status="clean",
                file_token=file_token,
                preview_analysis=analysis
            )

    except Exception as e:
        logger.error(f"Scan failed: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/generate", response_model=GenerateResult)
async def generate_config(request: GenerateRequest):
    """
    Step 2: Generate final config using verified/mapped headers.
    """
    temp_dir = sys_config.temp_uploads_dir
    file_path = temp_dir / request.file_token
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File token expired or invalid. Please re-scan.")
        
    try:
        # 1. Update Mapping Config TEMPORARILY or PERMANENTLY? 
        # For now, let's update strict mapping config so Scanner sees it.
        # But safer is to pass mappings DIRECTLY to Orchestrator -> Generator -> Scanner.
        
        # We need to enhance Orchestrator to accept 'manual_mappings' override
        orchestrator = Orchestrator()
        
        # TODO: This requires updating Orchestrator.generate_blueprint_bundle signature
        # to accept 'runtime_mappings' dict.
        
        result_path = orchestrator.generate_blueprint_bundle(
            template_path=file_path,
            output_dir=sys_config.bundled_dir,
            custom_prefix=request.customer_code,
            # User mappings passed here!
            runtime_mappings=request.mappings 
        )
        
        return GenerateResult(
            status="success",
            config_path=str(result_path),
            template_path=str(result_path.with_name(f"{request.customer_code}.xlsx")),
            message=f"Blueprint generated for {request.customer_code}"
        )

    except Exception as e:
        logger.error(f"Generation failed: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

class StringUtils:
    @staticmethod
    def is_unknown_col_id(col_id: str) -> bool:
        return col_id and col_id.startswith("col_unknown")

# --- Helper ---
def _format_label(col_id: str) -> str:
    """col_qty_pcs -> Qty Pcs"""
    return col_id.replace("col_", "").replace("_", " ").title()

@router.get("/options")
async def get_mapping_options():
    """
    Return list of valid system columns for mapping.
    Frontend uses this to populate the dropdown.
    """
    from core.blueprint_generator.rules import BlueprintRules
    
    options = []
    # Sort by ID or Priority? valid columns are in BlueprintRules.COLUMNS
    sorted_cols = sorted(BlueprintRules.COLUMNS.values(), key=lambda c: c.id)
    
    for col in sorted_cols:
        options.append({
            "id": col.id,
            "label": _format_label(col.id),
            "description": f"Internal ID: {col.id}" 
        })
        
    return options
