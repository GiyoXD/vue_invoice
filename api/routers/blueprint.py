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
    locale: str = "KH"
    mappings: Dict[str, str] = {} # {"Unknown Header": "col_remark"}
    footer_mappings: List[str] = []
    pricing_mode: str = "standard"  # 'standard' or 'net'

class GenerateResult(BaseModel):
    status: str
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
        # Save newly confirmed header and footer labels to global Config permanently
        if request.mappings or request.footer_mappings:
            from core.database.db_manager import get_global_mapping_config, save_global_mapping_config
            data = get_global_mapping_config()
            
            updated = False
            
            if request.footer_mappings:
                if "footer_label_mappings" not in data:
                    data["footer_label_mappings"] = {"keywords": []}
                existing_footers = data["footer_label_mappings"].get("keywords", [])
                for fm in request.footer_mappings:
                    if fm not in existing_footers:
                        existing_footers.append(fm)
                data["footer_label_mappings"]["keywords"] = existing_footers
                updated = True
                
            if request.mappings:
                filtered_mappings = {k: v for k, v in request.mappings.items() if v and v != "col_unknown"}
                if filtered_mappings:
                    if "header_text_mappings" not in data:
                        data["header_text_mappings"] = {"mappings": {}}
                    if "mappings" not in data["header_text_mappings"]:
                        data["header_text_mappings"]["mappings"] = {}
                    data["header_text_mappings"]["mappings"].update(filtered_mappings)
                    updated = True
            
            if updated:
                save_global_mapping_config(data)
                     
        # Run Generator
        orchestrator = Orchestrator()
        
        customer_code = request.customer_code
        locale = request.locale
        
        from core.database.db_manager import SessionLocal
        from core.database.repositories import BlueprintRepository
        db = SessionLocal()
        existing_template_json = None
        try:
            repo = BlueprintRepository(db)
            existing = repo.get_blueprint(customer_code, locale)
            if existing:
                existing_template_json = json.loads(existing.template_json)
        finally:
            db.close()
        
        config_data, template_json_data, template_xlsx_bytes = orchestrator.generate_blueprint_bundle(
            template_path=file_path,
            custom_prefix=customer_code,
            runtime_mappings=request.mappings,
            pricing_mode=request.pricing_mode,
            in_memory=True,
            existing_template_json=existing_template_json
        )
        
        # Save to SQLite directly in-memory
        from api.routers.templates import save_blueprint_to_db
        save_blueprint_to_db(
            customer_code=customer_code,
            locale=locale,
            config_data=config_data,
            template_json_data=template_json_data,
            xlsx_bytes=template_xlsx_bytes,
            filename=f"{customer_code}_{locale}.xlsx"
        )
        
        return GenerateResult(
            status="success",
            message=f"Blueprint generated and saved to database for {customer_code}"
        )

    except Exception as e:
        logger.error(f"Blueprint generation failed: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        # Clean up the uploaded Excel file now that the blueprint is generated
        try:
            if file_path.exists():
                file_path.unlink()
        except Exception as cleanup_err:
            logger.warning(f"Failed to delete temporary blueprint file {file_path}: {cleanup_err}")

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
    from core.blueprint_generator.schema import BlueprintSchema
    
    options = []
    # Sort by ID or Priority? valid columns are in BlueprintSchema.COLUMNS
    sorted_cols = sorted(BlueprintSchema.COLUMNS.values(), key=lambda c: c.id)
    
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
        from core.database.db_manager import get_global_mapping_config
        data = get_global_mapping_config()

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
        elif mapping_type == "sheet_classifications":
            flat = {}
            for s in data.get("aggregation_sheets", []):
                flat[s] = "aggregation"
            for s in data.get("processed_tables_sheets", []):
                flat[s] = "processed_tables"
            return flat
        elif mapping_type == "sheet_mappings":
            from core.database.db_manager import SessionLocal, GlobalMapSheet
            db = SessionLocal()
            try:
                sheets = db.query(GlobalMapSheet).all()
                res_dict = {}
                for s in sheets:
                    res_dict[s.sheet_name] = s.processing_type
                return res_dict
            finally:
                db.close()

        return data.get(mapping_type, {}).get("mappings", {})
    except Exception as e:
        logger.error(f"Failed to get mappings: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

class MappingsUpdateRequest(BaseModel):
    mapping_type: str = "header_text_mappings"
    mappings: Dict[str, Any]

@router.post("/mappings")
async def update_mappings(request: MappingsUpdateRequest):
    """
    Overwrite the specified global mapping dictionary.

    For column_definitions, the mappings dict is {col_id: "kw1, kw2, ..."}
    and is converted back to the structured format on save.
    """
    try:
        from core.database.db_manager import get_global_mapping_config, save_global_mapping_config
        data = get_global_mapping_config()

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
        elif request.mapping_type == "sheet_classifications":
            agg_sheets = []
            proc_sheets = []
            for name, p_type in request.mappings.items():
                if p_type == "processed_tables":
                    proc_sheets.append(name)
                else:
                    agg_sheets.append(name)
            data["aggregation_sheets"] = agg_sheets
            data["processed_tables_sheets"] = proc_sheets
        elif request.mapping_type == "sheet_mappings":
            agg_sheets = []
            proc_sheets = []
            for sheet_name, processing_type in request.mappings.items():
                if processing_type == "processed_tables":
                    proc_sheets.append(sheet_name)
                else:
                    agg_sheets.append(sheet_name)
            data["sheet_name_mappings"] = {"mappings": {}}
            data["aggregation_sheets"] = sorted(list(set(agg_sheets)))
            data["processed_tables_sheets"] = sorted(list(set(proc_sheets)))
        else:
            if request.mapping_type not in data:
                data[request.mapping_type] = {"mappings": {}}
            data[request.mapping_type]["mappings"] = request.mappings

        save_global_mapping_config(data)

        # Reload the mappings dynamically so the server doesn't need to be restarted
        try:
            from core.data_parser.config import load_and_update_mappings
            load_and_update_mappings()
            # Rebuild sheet_parser's pre-computed alias lookup after config change
            from core.data_parser.sheet_parser import _build_alias_lookup, _ALIAS_REVERSE_LOOKUP
            import core.data_parser.sheet_parser as _sp_module
            _sp_module._ALIAS_REVERSE_LOOKUP = _build_alias_lookup()
            
            from core.blueprint_generator.schema import BlueprintSchema
            BlueprintSchema.load_dynamic_columns(data)
        except Exception as e:
            logger.warning(f"Could not automatically reload mappings: {e}")

        return {"status": "success"}
    except Exception as e:
        logger.error(f"Failed to update mappings: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})
