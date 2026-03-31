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
    unconfirmed_footers: List[str] = []
    warnings: List[str] = []
    preview_analysis: Optional[Dict[str, Any]] = None

class GenerateRequest(BaseModel):
    file_token: str
    customer_code: str # e.g. "CLW"
    mappings: Dict[str, str] = {} # {"Unknown Header": "col_remark"}
    footer_mappings: List[str] = []

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
        unconfirmed_footers = []
        for sheet in analysis.get("sheets", []):
            for header in sheet.get("header_positions", []):
                if StringUtils.is_unknown_col_id(header.get("col_id")):
                    unknown_headers.append(header.get("keyword"))
            uf = sheet.get("unconfirmed_footer")
            if uf:
                unconfirmed_footers.append(uf)
        
        # Deduplicate
        unknown_headers = list(set(unknown_headers))
        unconfirmed_footers = list(set(unconfirmed_footers))
        
        if unknown_headers or unconfirmed_footers:
            return ScanResult(
                status="needs_mapping",
                file_token=file_token,
                unknown_headers=unknown_headers,
                unconfirmed_footers=unconfirmed_footers,
                warnings=analysis.get("warnings", []),
                preview_analysis=analysis
            )
        else:
             return ScanResult(
                status="clean",
                file_token=file_token,
                warnings=analysis.get("warnings", []),
                preview_analysis=analysis,
                unconfirmed_footers=[]
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
        # Save newly confirmed footer labels to global Config permanently
        if request.footer_mappings:
            config_path = sys_config.mapping_config_path
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if "footer_label_mappings" not in data:
                    data["footer_label_mappings"] = {"keywords": []}
                
                existing_footers = data["footer_label_mappings"].get("keywords", [])
                for fm in request.footer_mappings:
                    if fm not in existing_footers:
                        existing_footers.append(fm)
                data["footer_label_mappings"]["keywords"] = existing_footers
                
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                    
        # Run Generator
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

@router.get("/mappings")
async def get_mappings(mapping_type: str = "header_text_mappings"):
    """
    Get the global mapping dictionary of the specified type.
    Options: header_text_mappings, sheet_name_mappings, shipping_header_map

    For shipping_header_map, returns a flat dict of {col_id: "kw1, kw2, ..."}
    so the frontend can use the same key-value editor UI.
    """
    try:
        from core.system_config import sys_config
        with open(sys_config.mapping_config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        if mapping_type == "shipping_header_map":
            # Flatten to {col_id: "kw1, kw2"} for the UI
            col_defs = data.get("shipping_header_map", {})
            flat = {}
            for col_id, props in col_defs.items():
                if isinstance(props, dict):
                    flat[col_id] = ", ".join(props.get("keywords", []))
            return flat
        elif mapping_type == "footer_label_mappings":
            keywords = data.get("footer_label_mappings", {}).get("keywords", [])
            return {kw: "Footer Keyword" for kw in keywords}

        return data.get(mapping_type, {}).get("mappings", {})
    except Exception as e:
        logger.error(f"Failed to get mappings: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

class MappingsUpdateRequest(BaseModel):
    mapping_type: str = "header_text_mappings"
    mappings: Dict[str, str]

@router.post("/mappings")
async def update_mappings(request: MappingsUpdateRequest):
    """
    Overwrite the specified global mapping dictionary.

    For column_definitions, the mappings dict is {col_id: "kw1, kw2, ..."}
    and is converted back to the structured format on save.
    """
    try:
        from core.system_config import sys_config
        config_path = sys_config.mapping_config_path

        data = {}
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

        if request.mapping_type == "shipping_header_map":
            # Unflatten from {col_id: "kw1, kw2"} back to structured format
            existing = data.get("shipping_header_map", {})
            for col_id, kw_str in request.mappings.items():
                keywords = [k.strip() for k in kw_str.split(",") if k.strip()]
                if col_id in existing and isinstance(existing[col_id], dict):
                    existing[col_id]["keywords"] = keywords
                else:
                    existing[col_id] = {"keywords": keywords, "format": "@"}
            data["shipping_header_map"] = existing
        elif request.mapping_type == "footer_label_mappings":
            existing = data.get("footer_label_mappings", {})
            existing["keywords"] = list(request.mappings.keys())
            data["footer_label_mappings"] = existing
        else:
            if request.mapping_type not in data:
                data[request.mapping_type] = {"mappings": {}}
            data[request.mapping_type]["mappings"] = request.mappings

        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

        return {"status": "success"}
    except Exception as e:
        logger.error(f"Failed to update mappings: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})
