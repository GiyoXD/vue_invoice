from fastapi import FastAPI, BackgroundTasks, UploadFile, File, Body
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
import shutil
import os
import json
from pathlib import Path
import uuid
from typing import List, Optional
import datetime

# Import core orchestrator
from core.orchestrator import Orchestrator
import subprocess
import sys

# Define Project Root
from core.system_config import sys_config
PROJECT_ROOT = Path(__file__).resolve().parent.parent

CONFIG_GEN_DIR = PROJECT_ROOT / "core" / "blueprint_generator"
MAPPING_CONFIG_PATH = sys_config.mapping_config_path

SYSTEM_HEADERS = [
    "col_po", "col_item", "col_desc", "col_qty_pcs", "col_qty_sf", 
    "col_unit_price", "col_amount", "col_net", "col_gross", "col_cbm", 
    "col_pallet", "col_remarks", "col_static", "col_dc"
]

app = FastAPI()

# Mount frontend
app.mount("/frontend", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="frontend")
app.mount("/static", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="static")

# Include Routers
from api.routers import blueprint
app.include_router(blueprint.router)


orchestrator = Orchestrator()

# Temporary storage for uploads
UPLOAD_DIR = Path("temp_uploads")
OUTPUT_DIR = Path("output")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

class GenerateRequest(BaseModel):
    identifier: str
    json_path: str
    invoice_no: str
    invoice_date: str
    invoice_ref: Optional[str] = ""

@app.get("/api/health")
async def health_check():
    return {"status": "ok"}

@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    """
    Uploads an Excel file and processes it to JSON.
    Returns the identifier, json path, and asset availability status.
    
    The asset_status field tells the frontend whether the required
    config and template files exist for invoice generation.
    """
    try:
        file_path = UPLOAD_DIR / file.filename
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Process to JSON using Orchestrator
        json_output_dir = UPLOAD_DIR / "processed"
        json_output_dir.mkdir(exist_ok=True)

        json_path, identifier = orchestrator.process_excel_to_json(file_path, json_output_dir)
        
        # Default Invoice No to filename stem
        default_inv_no = Path(file.filename).stem
        
        # === CHECK ASSET AVAILABILITY ===
        # Use the InvoiceAssetResolver to see if we have config/template for this file
        from core.invoice_generator.resolvers import InvoiceAssetResolver
        
        resolver = InvoiceAssetResolver(
            base_config_dir=sys_config.registry_dir,
            base_template_dir=sys_config.templates_dir
        )
        
        # Check if assets can be resolved for this input
        assets = resolver.resolve_assets_for_input_file(str(json_path))
        
        asset_status = {
            "ready": assets is not None,
            "config_found": False,
            "template_found": False,
            "config_path": None,
            "template_path": None,
            "bundled_dir": str(sys_config.bundled_dir),
            "message": ""
        }
        
        if assets:
            asset_status["config_found"] = True
            asset_status["template_found"] = True
            asset_status["config_path"] = str(assets.config_path)
            asset_status["template_path"] = str(assets.template_path)
            asset_status["message"] = "Ready to generate invoice."
        else:
            # Provide helpful guidance
            prefix = identifier[:2] if len(identifier) >= 2 else identifier
            asset_status["message"] = (
                f"No config/template found for '{identifier}'. "
                f"Expected: bundled/{prefix}/ folder with {prefix}_config.json and {prefix}.xlsx"
            )
        
        return {
            "status": "success",
            "file_name": file.filename,
            "identifier": identifier,
            "json_path": str(json_path),
            "default_inv_no": default_inv_no,
            "asset_status": asset_status,
            "message": "File processed successfully"
        }
    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={
            "error": str(e), 
            "traceback": traceback.format_exc(),
            "step": "Upload & Parse"
        })




@app.post("/api/generate")
async def generate_invoice(request: GenerateRequest):
    """
    Trigger invoice generation with metadata overrides.
    """
    try:
        # Resolve paths
        json_path_obj = Path(request.json_path)
        if not json_path_obj.exists():
             return JSONResponse(status_code=404, content={"error": "JSON file not found. Please upload again."})

        # Define output path
        output_path = OUTPUT_DIR / request.identifier
        output_path.mkdir(exist_ok=True)

        # Default Template/Config dirs - use bundled directory from config
        template_dir = sys_config.bundled_dir
        config_dir = sys_config.bundled_dir

        # Load the existing JSON data
        try:
            with open(json_path_obj, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
        except Exception as e:
             return JSONResponse(status_code=500, content={"error": f"Failed to load JSON data: {str(e)}"})

        # Update with overrides
        if "invoice_info" not in full_data:
            full_data["invoice_info"] = {}
        
        full_data["invoice_info"]["col_inv_no"] = request.invoice_no
        full_data["invoice_info"]["col_inv_date"] = request.invoice_date
        full_data["invoice_info"]["col_inv_ref"] = request.invoice_ref
        
        # Pass FULL modified data to orchestrator
        result_path = orchestrator.generate_invoice(
            json_path=json_path_obj,
            output_path=output_path / f"{request.identifier}_Invoice.xlsx",
            template_dir=template_dir,
            config_dir=config_dir,
            input_data_dict=full_data 
        )
        
        # Open file explorer to the output directory
        try:
             os.startfile(result_path.parent)
        except Exception:
             pass # Ignore if fails (e.g. headless)

        # Read the generated metadata to send back validation info
        metadata_content = {}
        try:
            meta_path = result_path.parent / f"{result_path.stem}_metadata.json"
            if meta_path.exists():
                with open(meta_path, 'r', encoding='utf-8') as f:
                    metadata_content = json.load(f)
        except Exception:
            pass

        return {
            "status": "completed",
            "output_path": str(result_path),
            "message": f"Invoice generated at {result_path}",
            "metadata": metadata_content
        }


    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={
            "error": str(e), 
            "traceback": traceback.format_exc(),
            "step": "Invoice Generation"
        })


@app.get("/api/history")
async def get_history():
    """
    Retrieve list of past runs from the run_log registry.
    """
    run_log_dir = Path("run_log")
    history = []
    
    if run_log_dir.exists():
        # List all json files
        for f in run_log_dir.glob("*_metadata.json"):
            try:
                # Parse filename: Timestamp_Filename_metadata.json
                # We can just read the file content for better info
                with open(f, 'r', encoding='utf-8') as json_file:
                    data = json.load(json_file)
                    
                    history.append({
                        "filename": f.name,
                        "timestamp": data.get("timestamp"),
                        "output_file": data.get("output_file"),
                        "status": data.get("status"),
                        "item_count": data.get("database_export", {}).get("summary", {}).get("item_count", 0),
                        "total_sqft": data.get("database_export", {}).get("summary", {}).get("total_sqft", 0)
                    })
            except Exception:
                continue # Skip corrupted files

    # Sort by timestamp descending
    history.sort(key=lambda x: x["timestamp"] or "", reverse=True)
    return history

@app.get("/api/history/view")
async def view_history_item(filename: str):
    """
    Retrieve specific metadata file content from run_log.
    """
    run_log_dir = Path("run_log")
    file_path = run_log_dir / filename
    
    if not file_path.exists():
        # Security check: ensure simple filename, no traversal
        if ".." in filename or "/" in filename:
             return JSONResponse(status_code=400, content={"error": "Invalid filename"})
        return JSONResponse(status_code=404, content={"error": "File not found"})
        
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": "Could not read file"})

# --- Helper Functions for Template Extractor ---

def get_header_suggestions(header_text: str) -> str:
    header_lower = header_text.lower()
    suggestions = {
        "col_po": ['p.o', 'po'], "col_item": ['item', 'no.'], "col_desc": ['description', 'desc'],
        "col_qty_sf": ['quantity', 'qty'], "col_unit_price": ['unit', 'price'], "col_amount": ['amount', 'total', 'value'],
        "col_net": ['n.w', 'net'], "col_gross": ['g.w', 'gross'], "col_cbm": ['cbm'],
        "col_pallet": ['pallet'], "col_remarks": ['remarks', 'notes'], "col_static": ['mark', 'note']
    }
    for col_id, keywords in suggestions.items():
        if any(word in header_lower for word in keywords):
            return col_id
    return "col_unknown"

def get_missing_headers(analysis_file_path: str):
    try:
        with open(analysis_file_path, 'r', encoding='utf-8') as f:
            analysis_data = json.load(f)
            
        missing_headers = []
        
        # Iterate through all sheets and headers
        for sheet in analysis_data.get('sheets', []):
            for header_pos in sheet.get('header_positions', []):
                col_id = header_pos.get('col_id', '')
                header_text = header_pos.get('keyword', '')
                
                # Logic: If ID is "col_unknown_...", it needs mapping.
                if col_id.startswith("col_unknown"):
                    # Double check if it's already in mapping (frontend might have missed it or partial reload)
                    # But generally, if Scanner said Unknown, it means it wasn't in User Mapping EITHER.
                    missing_headers.append({
                        "text": header_text, 
                        "suggestion": get_header_suggestions(header_text)
                    })
        
        return missing_headers
    except Exception:
        import traceback
        traceback.print_exc()
        return []

def update_mapping_config(new_mappings: dict):
    try:
        mapping_data = {"header_text_mappings": {"mappings": {}}}
        if MAPPING_CONFIG_PATH.exists():
            with open(MAPPING_CONFIG_PATH, 'r', encoding='utf-8') as f:
                mapping_data = json.load(f)
        
        if "header_text_mappings" not in mapping_data: mapping_data["header_text_mappings"] = {"mappings": {}}
        if "mappings" not in mapping_data["header_text_mappings"]: mapping_data["header_text_mappings"]["mappings"] = {}

        # Filter out 'col_unknown' or invalid mappings
        filtered_mappings = {k: v for k, v in new_mappings.items() if v and v != "col_unknown"}

        if filtered_mappings:
            mapping_data["header_text_mappings"]["mappings"].update(filtered_mappings)

            with open(MAPPING_CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(mapping_data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

@app.get("/api/download")
async def download_file(path: str):
    """
    Download a file from an absolute path.
    """
    file_path = Path(path)
    if not file_path.exists() or not file_path.is_file():
        return JSONResponse(status_code=404, content={"error": "File not found"})
        
    return FileResponse(file_path, filename=file_path.name, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# --- Template Extractor API ---
from core.system_config import sys_config

TEMP_DIR = sys_config.temp_uploads_dir
# Note: Template generation now outputs to bundled/{prefix}/ folder via sys_config.bundled_dir


class TemplateConfig(BaseModel):
    file_prefix: str
    user_mappings: dict
    temp_filename: str

@app.post("/api/template/analyze")
async def analyze_template(file: UploadFile = File(...)):
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = TEMP_DIR / file.filename
    
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        analysis_output_path = TEMP_DIR / f"{file.filename}_analysis.json"
        
        # Use Orchestrator
        json_output = orchestrator.analyze_template(temp_path, legacy_format=True)
        
        # Save to analysis_output_path for get_missing_headers to read
        with open(analysis_output_path, 'w', encoding='utf-8') as f:
            f.write(json_output)
            
        missing_headers = get_missing_headers(str(analysis_output_path))
        
        # Clean up analysis file immediately, keep excel for next step
        if analysis_output_path.exists():
            analysis_output_path.unlink()
            
        return {
            "missing_headers": missing_headers,
            "temp_filename": file.filename,
            "suggested_prefix": file.filename.split('.')[0]
        }
    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={
            "error": str(e), 
            "traceback": traceback.format_exc(),
            "step": "Template Analysis"
        })

@app.post("/api/template/generate")
async def generate_template(config: TemplateConfig):
    """
    Generate a template bundle for a new customer.
    
    Creates a bundled folder structure:
    bundled/{prefix}/
      - {prefix}.xlsx (template)
      - {prefix}_config.json (config)
    """
    try:
        # 1. Update global mappings if user provided any
        if config.user_mappings:
            if not update_mapping_config(config.user_mappings):
                return JSONResponse(status_code=500, content={"error": "Failed to update mapping config"})

        # 2. Setup paths - NEW BUNDLED STRUCTURE
        temp_path = TEMP_DIR / config.temp_filename
        if not temp_path.exists():
            return JSONResponse(status_code=404, content={"error": "Original uploaded file not found. Please re-upload."})

        # 3. Run Generator via Orchestrator
        result_path = orchestrator.generate_blueprint_bundle(
            template_path=temp_path,
            output_dir=sys_config.bundled_dir,
            custom_prefix=config.file_prefix
        )
        
        if not result_path:
            return JSONResponse(status_code=500, content={"error": "Config generation failed (no result path returned)"})

        # The generator creates the folder using the custom_prefix
        customer_bundle_dir = result_path.parent

        return {
            "status": "success", 
            "message": f"Template {config.file_prefix} created successfully!",
            "bundle_path": str(customer_bundle_dir)
        }

    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={
            "error": str(e), 
            "traceback": traceback.format_exc(),
            "step": "Template Generation"
        })




# --- Template Inspector API ---

@app.get("/api/templates")
async def list_templates():
    """
    List all available template bundles that have a generated JSON template.
    """
    bundled_dir = sys_config.bundled_dir
    templates = []
    
    if bundled_dir.exists():
        for item in bundled_dir.iterdir():
            if item.is_dir():
                # Check for {DirectoryName}_template.json
                template_json_path = item / f"{item.name}_template.json"
                if template_json_path.exists():
                     # Get creation time
                    try:
                        stats = template_json_path.stat()
                        # read fingerprint for source file
                        source_file = "Unknown"
                        try:
                            with open(template_json_path, 'r', encoding='utf-8') as f:
                                # Read first few lines or parse partly to avoid loading huge file? 
                                # Actually json.load is fast enough for metadata if file isn't massive.
                                # But let's just use json.load for now.
                                data = json.load(f)
                                source_file = data.get("fingerprint", {}).get("source_file", "Unknown")
                        except Exception:
                            pass

                        templates.append({
                            "name": item.name,
                            "path": str(template_json_path),
                            "modified": datetime.datetime.fromtimestamp(stats.st_mtime).isoformat(),
                            "source_file": source_file
                        })
                    except Exception:
                        pass
                        
    return templates

@app.get("/api/template/view")
async def view_template(name: str):
    """
    Get the content of a specific template JSON.
    """
    bundled_dir = sys_config.bundled_dir
    # Sanitize name to prevent traversal
    safe_name = Path(name).name
    template_path = bundled_dir / safe_name / f"{safe_name}_template.json"
    
    if not template_path.exists():
        return JSONResponse(status_code=404, content={"error": "Template JSON not found"})
        
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to read template: {str(e)}"})
