import logging
from fastapi import FastAPI, BackgroundTasks, UploadFile, File, Body
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, FileResponse, StreamingResponse, RedirectResponse
from pydantic import BaseModel
import shutil
import os
import json
import io
import csv
from pathlib import Path
import uuid
from typing import List, Optional, Any
import datetime
from fastapi import Depends
from sqlalchemy.orm import Session
from core.database import db_manager
from core.database.db_manager import get_db, ProcessedData, InvoiceItem, init_db, engine, Base, get_cambodia_time

# Initialize logging FIRST before any other core imports
from core.system_config import sys_config
from core.logger_config import setup_logging
setup_logging(log_dir=sys_config.run_log_dir)

# Import core orchestrator
from core.orchestrator import Orchestrator
from core.data_parser.data_processor import DataValidationError
import subprocess
import sys

# Define Project Root
PROJECT_ROOT = Path(__file__).resolve().parent.parent

CONFIG_GEN_DIR = PROJECT_ROOT / "core" / "blueprint_generator"
MAPPING_CONFIG_PATH = sys_config.mapping_config_path

SYSTEM_HEADERS = [
    "col_po", "col_item", "col_desc", "col_qty_pcs", "col_qty_sf", 
    "col_unit_price", "col_amount", "col_net", "col_gross", "col_cbm", 
    "col_pallet", "col_remarks", "col_static", "col_dc", "col_hs_code"
]

app = FastAPI()
logger = logging.getLogger(__name__)

# Mount frontend
app.mount("/frontend", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="frontend")
app.mount("/static", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="static")

# Include Routers
from api.routers import blueprint
app.include_router(blueprint.router)

@app.get("/")
def redirect_to_frontend():
    return RedirectResponse(url="/frontend/")


# Initialize Database
db_manager.init_db()


orchestrator = Orchestrator()

# Temporary storage for uploads
UPLOAD_DIR = sys_config.temp_uploads_dir
OUTPUT_DIR = sys_config.output_dir
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

class GenerateRequest(BaseModel):
    identifier: str
    json_path: str
    invoice_no: str
    invoice_date: str
    invoice_ref: Optional[str] = ""
    generate_standard: bool = True
    generate_custom: bool = False
    generate_daf: bool = False
    generate_kh: bool = False
    generate_vn: bool = False
    price_adjustment: Optional[List[List[Any]]] = None

@app.get("/api/health")
async def health_check():
    return {"status": "ok"}

@app.post("/api/upload")
def upload_excel(file: UploadFile = File(...)):
    """
    Uploads an Excel file and processes it to JSON.
    Returns the identifier, json path, and asset availability status.
    
    The asset_status field tells the frontend whether the required
    config and template files exist for invoice generation.
    """
    try:
        print(f"DEBUG: Received upload request for {file.filename}") # Direct console log
        
        # Read the file into memory
        file_bytes = file.file.read()
        buffer = io.BytesIO(file_bytes)
            
        # Process to JSON using Orchestrator
        json_output_dir = UPLOAD_DIR / "processed"
        json_output_dir.mkdir(exist_ok=True)

        json_path, identifier = orchestrator.process_excel_to_json(
            buffer, 
            json_output_dir,
            input_filename_override=file.filename
        )
        
        # Default Invoice No to filename stem
        default_inv_no = Path(file.filename).stem
        
        # === CHECK ASSET AVAILABILITY ===
        from core.invoice_generator.resolvers import InvoiceAssetResolver
        
        resolver = InvoiceAssetResolver(
            base_config_dir=sys_config.registry_dir,
            base_template_dir=sys_config.templates_dir
        )
        
        # Check if assets can be resolved for this input
        assets = resolver.resolve_assets_for_input_file(str(json_path))
        
        # Check for KH/VN variants
        variants = resolver.resolve_all_variants(str(json_path))
        
        asset_status = {
            "ready": assets is not None,
            "config_found": False,
            "template_found": False,
            "config_path": None,
            "template_path": None,
            "bundled_dir": str(sys_config.bundled_dir),
            "message": "",
            "variants": []
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
        
        # --- Read warnings from generated JSON ---
        warnings_list = []
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                parsed_data = json.load(f)
                warnings_list = parsed_data.get('metadata', {}).get('warnings', [])
        except Exception as e:
            logger.warning(f"Could not read warnings from JSON output: {e}")

        # Add variant info
        if variants:
            asset_status["variants"] = [
                {
                    "suffix": v["suffix"],
                    "config_path": str(v["config_path"]),
                    "template_path": str(v["template_path"])
                }
                for v in variants
            ]
        
        return {
            "status": "success",
            "file_name": file.filename,
            "identifier": identifier,
            "json_path": str(json_path),
            "default_inv_no": default_inv_no,
            "warnings": warnings_list,
            "asset_status": asset_status,
            "message": "File processed successfully"
        }
    except DataValidationError as ve:
        # User-facing validation errors get a clean response (no traceback noise)
        return JSONResponse(status_code=422, content={
            "error": str(ve),
            "step": "Data Validation"
        })
    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={
            "error": str(e), 
            "traceback": traceback.format_exc(),
            "step": "Upload & Parse"
        })




@app.post("/api/generate")
def generate_invoice(request: GenerateRequest):
    """
    Trigger invoice generation with metadata overrides.
    Supports generating multiple variations (Standard, Custom, DAF).
    """
    try:
        # Resolve paths
        json_path_obj = Path(request.json_path)
        if not json_path_obj.exists():
             return JSONResponse(status_code=404, content={"error": "JSON file not found. Please upload again."})

        # Define base output path (not created on disk to avoid clutter)
        base_output_dir = OUTPUT_DIR / request.identifier

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

        # Price adjustments (list of [description, value] pairs)
        from core.invoice_generator.utils.aggregation_modifier import apply_aggregation_adjustment
        price_adj = request.price_adjustment
        if price_adj:
            full_data = apply_aggregation_adjustment(full_data, price_adj)
        
        # PERSIST: Save the updated JSON back to disk so the user can see the changes
        try:
            with open(json_path_obj, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, indent=4)
        except Exception as e:
            logger.warning(f"Failed to persist updated JSON to {json_path_obj}: {e}")
        
        results = []
        errors = []
        generated_files = []
        primary_metadata = {}
        processed_any = False

        # Define mode tasks
        mode_tasks = []
        if request.generate_standard:
            mode_tasks.append({
                "suffix": "", 
                "flags": [],
                "name": "Standard Invoice"
            })
        if request.generate_custom:
            mode_tasks.append({
                "suffix": "_Custom", 
                "flags": ["--custom"],
                "name": "Custom Invoice"
            })
        if request.generate_daf:
            mode_tasks.append({
                "suffix": "_DAF", 
                "flags": ["--DAF"],
                "name": "DAF Invoice"
            })

        if not mode_tasks:
             return JSONResponse(status_code=400, content={"error": "No invoice type selected."})

        # Determine variant tasks (KH/VN or default)
        variant_tasks = []
        if request.generate_kh or request.generate_vn:
            # Resolve variants from the bundle folder
            from core.invoice_generator.resolvers import InvoiceAssetResolver
            resolver = InvoiceAssetResolver(
                base_config_dir=sys_config.registry_dir,
                base_template_dir=sys_config.templates_dir
            )
            all_variants = resolver.resolve_all_variants(str(json_path_obj))
            variant_map = {v["suffix"]: v for v in all_variants}
            
            if request.generate_kh and "_KH" in variant_map:
                variant_tasks.append(variant_map["_KH"])
            if request.generate_vn and "_VN" in variant_map:
                variant_tasks.append(variant_map["_VN"])
        
        # If no variants selected, run with default (no variant suffix, no explicit paths)
        if not variant_tasks:
            variant_tasks = [{"suffix": "", "config_path": None, "template_path": None}]

        # Cross-product: variant × mode
        for variant in variant_tasks:
            variant_suffix = variant["suffix"]  # e.g. "_KH" or ""
            
            for task in mode_tasks:
                try:
                    # Build filename: IDENTIFIER_Invoice_KH.xlsx or IDENTIFIER_Invoice_KH_Custom.xlsx
                    filename = f"{request.identifier}_Invoice{variant_suffix}{task['suffix']}.xlsx"
                    output_path = base_output_dir / filename

                    # Build flags with explicit config/template if variant has them
                    flags = list(task['flags'])
                    if variant.get("config_path"):
                        flags.extend(["--config", str(variant["config_path"])])
                    if variant.get("template_path"):
                        flags.extend(["--template", str(variant["template_path"])])

                    # Generate
                    result = orchestrator.generate_invoice(
                        json_path=json_path_obj,
                        output_path=output_path,
                        template_dir=template_dir,
                        config_dir=config_dir,
                        input_data_dict=full_data,
                        flags=flags,
                        return_bytes=True
                    )
                    
                    if result:
                        filename, file_bytes = result
                        results.append(filename)
                        # Store in memory temporarily
                        generated_files.append((filename, file_bytes))
                        processed_any = True

                    # Try to capture metadata from the first successful generation session if available
                    # Actually wait, GenerationSession writes to output_path parent dir which we just overrode if we didn't save?
                    
                except Exception as e:
                    import traceback
                    task_name = f"{variant_suffix.lstrip('_')} {task['name']}" if variant_suffix else task['name']
                    error_msg = f"Failed to generate {task_name}: {str(e)}"
                    print(traceback.format_exc())
                    errors.append(error_msg)

        if not processed_any and errors:
             # All failed
             return JSONResponse(status_code=500, content={
                "error": "All generation tasks failed.", 
                "details": errors
            })

        # Gather metadata from the database or run_log if needed?
        # Actually metadata is stored during extract, we can just return what we have or nothing
        # The user's main requirement is that the file is sent back via web ram.

        final_payload_files = []
        if generated_files:
            import zipfile
            import io
            import base64
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in generated_files:
                    zf.writestr(fname, fbytes)
            
            zip_buffer.seek(0)
            zip_b64 = base64.b64encode(zip_buffer.read()).decode('utf-8')
            zip_name = f"Invoices_{request.identifier}.zip"
            
            final_payload_files.append({
                "filename": zip_name,
                "mime_type": "application/zip",
                "content": zip_b64
            })

        return {
            "status": "completed",
            "output_paths": results,
            "message": f"Generated {len(results)} invoices.",
            "metadata": primary_metadata,
            "files": final_payload_files,
            "errors": errors if errors else None
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
    Retrieve list of past runs from the run_log registry 
    AND the processed directory for easy inspection.
    """
    history = []
    
    # 1. Check Run Logs (Historical metadata)
    run_log_dir = Path("run_log")
    if run_log_dir.exists():
        for f in run_log_dir.glob("*_metadata.json"):
            try:
                # Avoid duplicates if already in DB (unlikely for now but safe)
                if any(h["filename"] == f.name for h in history): continue
                
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
            except Exception:
                continue

    # 3. Check Processed Dir (Intermediate JSONs - PENDING)
    processed_dir = sys_config.temp_uploads_dir / "processed"
    if processed_dir.exists():
        for f in processed_dir.glob("*.json"):
            try:
                # IMPORTANT: If it's already in the "accepted" history (database),
                # we don't show it twice as "processed"
                if any(h["filename"] == f.name for h in history):
                    continue
                
                stats = f.stat()
                with open(f, 'r', encoding='utf-8') as json_file:
                    data = json.load(json_file)
                    
                    # Try to find item count/sqft if it's the new multi_table format
                    item_count = 0
                    if "multi_table" in data and data["multi_table"]:
                        item_count = sum(len(table) for table in data["multi_table"] if isinstance(table, list))
                    
                    history.append({
                        "filename": f.name,
                        "type": "processed", # This means "Pending"
                        "timestamp": datetime.datetime.fromtimestamp(stats.st_mtime).isoformat(),
                        "output_file": f.name,
                        "status": "Ready",
                        "item_count": item_count,
                        "total_sqft": 0 # Square footage not always available at top level for raw output
                    })
            except Exception:
                continue

    # Sort by timestamp descending
    history.sort(key=lambda x: x["timestamp"] or "", reverse=True)
    return history

@app.get("/api/history/view")
async def view_history_item(filename: str, source: str = "run_log"):
    """
    Retrieve specific JSON content from run_log or processed dir.
    """
    if ".." in filename or "/" in filename or "\\" in filename:
         return JSONResponse(status_code=400, content={"error": "Invalid filename"})

    if source == "processed":
        file_path = sys_config.temp_uploads_dir / "processed" / filename
    elif source == "accepted":
        # For accepted runs, retrieve from DB data_payload
        try:
            db = next(get_db())
            record = db.query(ProcessedData).filter(ProcessedData.filename == filename).first()
            if record:
                return record.data_payload
            return JSONResponse(status_code=404, content={"error": "Database record not found"})
        except Exception as e:
            return JSONResponse(status_code=500, content={"error": str(e)})
    else:
        file_path = Path("run_log") / filename
        # Robustness: If not found in run_log, check processed as well
        if not file_path.exists():
            fallback_path = sys_config.temp_uploads_dir / "processed" / filename
            if fallback_path.exists():
                file_path = fallback_path
    
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"error": "File not found"})
        
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Could not read file: {str(e)}"})

# --- Registry Management ---

class AcceptRejectRequest(BaseModel):
    filename: str

@app.post("/api/registry/check")
async def check_invoice_exists(req: AcceptRejectRequest, db: Session = Depends(get_db)):
    """
    Checks if a processed JSON file has already been accepted and exists in the database.
    """
    existing = db.query(ProcessedData).filter(ProcessedData.filename == req.filename).first()
    return {"exists": existing is not None}

@app.post("/api/registry/accept")
async def accept_invoice(req: AcceptRejectRequest, db: Session = Depends(get_db)):
    """
    Accepts a processed JSON file:
    1. Reads full data from disk.
    2. Calculates summary (item_count, sqft, amount).
    3. Saves everything to SQLite 'processed_data' table.
    4. Deletes the temporary file.
    """
    processed_dir = sys_config.temp_uploads_dir / "processed"
    file_path = processed_dir / req.filename
    
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"error": "Pending file not found"})
        
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            
        # Delete existing items immediately to prepare for new ones
        db.query(InvoiceItem).filter(InvoiceItem.invoice_id == req.filename).delete()
        
        items_to_add = []
        item_count = 0
        total_sqft = 0.0
        total_amount = 0.0
        # Extract total_pallets securely from footer_data.grand_total to avoid duplicate summation
        total_pallets = 0.0
        try:
            footer_data = data.get("footer_data", {})
            grand_total = footer_data.get("grand_total", {})
            gp_count = grand_total.get("col_pallet_count", 0)
            total_pallets = float(gp_count) if gp_count else 0.0
        except Exception:
            total_pallets = 0.0
        
        source_tables = data.get("raw_data") or data.get("multi_table") or []
        for table in source_tables:
            if isinstance(table, list):
                for row in table:
                    # Extract fields safely, handling type conversions
                    qty_pcs = row.get("col_qty_pcs")
                    qty_sf = row.get("col_qty_sf")
                    pallet_count = row.get("col_pallet_count")
                    net = row.get("col_net")
                    gross = row.get("col_gross")
                    unit_price = row.get("col_unit_price")
                    amount = row.get("col_amount")

                    def to_float(val):
                        try:
                            if isinstance(val, str):
                                val = val.replace(',', '').replace('$', '').strip()
                            return float(val) if val is not None and str(val).strip() != "" else 0.0
                        except: return 0.0

                    sqft_val = to_float(qty_sf)
                    amount_val = to_float(amount)
                    pallet_val = to_float(pallet_count)
                    item_count += 1
                    total_sqft += sqft_val
                    total_amount += amount_val
                    # Note: We NO LONGER sum total_pallets here since it inflates when items share pallets

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
                        col_qty_pcs=to_float(qty_pcs),
                        col_qty_sf=sqft_val,
                        col_pallet_count=pallet_val,
                        col_pallet_count_raw=str(row.get("col_pallet_count_raw", "")),
                        col_net=to_float(net),
                        col_gross=to_float(gross),
                        col_cbm_raw=str(row.get("col_cbm_raw", row.get("col_cbm", ""))),
                        col_hs_code=str(row.get("col_hs_code", "")),
                        col_unit_price=to_float(unit_price),
                        col_amount=amount_val,
                        is_adjustment=0,
                        timestamp=get_cambodia_time()
                    )
                    items_to_add.append(item)

        # 3. Extract price_adjustments
        if "price_adjustment" in data:
            for adj in data["price_adjustment"]:
                 def to_float(val):
                     try: return float(val) if val is not None and str(val).strip() != "" else 0.0
                     except: return 0.0
                 amount = adj.get("amount", 0.0)
                 amount_val = to_float(amount)
                 total_amount += amount_val
                 
                 item = InvoiceItem(
                     invoice_id=req.filename,
                     col_desc=str(adj.get("description", "")),
                     col_amount=amount_val,
                     is_adjustment=1,
                     timestamp=get_cambodia_time()
                 )
                 items_to_add.append(item)
                 
        if items_to_add:
            db.add_all(items_to_add)

        # Update ProcessedData
        existing = db.query(ProcessedData).filter(ProcessedData.filename == req.filename).first()
        if existing:
            existing.item_count = item_count
            existing.total_sqft = total_sqft
            existing.total_amount = total_amount
            existing.total_pallets = total_pallets
            existing.data_payload = data
            existing.timestamp = get_cambodia_time()
        else:
            new_record = ProcessedData(
                filename=req.filename,
                item_count=item_count,
                total_sqft=total_sqft,
                total_amount=total_amount,
                total_pallets=total_pallets,
                data_payload=data
            )
            db.add(new_record)

        db.commit()
        
        # Delete the file after successful DB save

        file_path.unlink()
        
        return {"status": "success", "message": f"Invoice {req.filename} accepted and saved to database."}
        
    except Exception as e:
        import traceback
        logger.error(traceback.format_exc())
        return JSONResponse(status_code=500, content={"error": f"Failed to accept invoice: {str(e)}"})

@app.post("/api/registry/reject")
async def reject_invoice(req: AcceptRejectRequest):
    """
    Rejects a processed JSON file by deleting it from disk.
    """
    processed_dir = sys_config.temp_uploads_dir / "processed"
    file_path = processed_dir / req.filename
    
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"error": "Pending file not found"})
        
    try:
        file_path.unlink()
        return {"status": "success", "message": f"Invoice {req.filename} rejected and deleted."}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to reject invoice: {str(e)}"})

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

@app.get("/api/registry/export")
async def export_registry(
    start_date: Optional[str] = None, 
    end_date: Optional[str] = None, 
    db: Session = Depends(get_db)
):
    """
    Exports the processed data registry as a CSV file within a given time interval.
    """
    try:
        query = db.query(InvoiceItem)
        
        if start_date:
            try:
                start_dt = datetime.datetime.fromisoformat(start_date)
                query = query.filter(InvoiceItem.timestamp >= start_dt)
            except ValueError:
                pass
                
        if end_date:
            try:
                # End date should include the whole day if just YYYY-MM-DD
                if len(end_date) <= 10:
                    end_dt = datetime.datetime.fromisoformat(end_date) + datetime.timedelta(days=1)
                else:
                    end_dt = datetime.datetime.fromisoformat(end_date)
                query = query.filter(InvoiceItem.timestamp < end_dt)
            except ValueError:
                pass

        results = query.order_by(InvoiceItem.timestamp.desc(), InvoiceItem.invoice_id).all()
        
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Write Header
        writer.writerow([
            "ID", "Invoice ID", "Timestamp", "DC", "PO", "Production Order No", 
            "Production Date", "Line No", "Direction", "Item", "Reference Code", 
            "Description", "Level", "Grade", "Qty (PCS)", "Qty (SF)", "Pallet Count", 
            "Pallet Count Raw", "Net Weight", "Gross Weight", "CBM Raw", 
            "HS Code", "Unit Price", "Amount", "Is Adjustment"
        ])
        
        # Write Rows
        total_sqft = 0.0
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
                row.col_pallet_count_raw,
                row.col_net,
                row.col_gross,
                row.col_cbm_raw,
                row.col_hs_code,
                row.col_unit_price,
                row.col_amount,
                "Yes" if row.is_adjustment else "No"
            ])
            
            total_sqft += float(row.col_qty_sf or 0.0)
            total_pallets += float(row.col_pallet_count or 0.0)
            total_amount += float(row.col_amount or 0.0)
            
        # Write Summary Row
        summary_row = [""] * 25
        summary_row[1] = "TOTAL"
        summary_row[15] = round(total_sqft, 2)
        summary_row[16] = round(total_pallets, 2)
        summary_row[23] = round(total_amount, 2)
        
        writer.writerow([])
        writer.writerow(summary_row)
            
        output.seek(0)
        csv_data = "\ufeff" + output.getvalue()
        
        filename = f"export_{get_cambodia_time().strftime('%Y%m%d_%H%M%S')}.csv"
        
        return StreamingResponse(
            iter([csv_data]),
            media_type="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        import traceback
        logger.error(traceback.format_exc())
        return JSONResponse(status_code=500, content={"error": f"Failed to export data: {str(e)}"})

@app.get("/api/registry/list")
async def list_registry(db: Session = Depends(get_db)):
    """
    Returns a list of recent processed invoices for preview.
    """
    try:
        results = db.query(ProcessedData).order_by(ProcessedData.timestamp.desc()).limit(10).all()
        return [
            {
                "id": row.id,
                "filename": row.filename,
                "timestamp": row.timestamp.isoformat(),
                "item_count": row.item_count,
                "total_amount": row.total_amount,
                "total_sqft": row.total_sqft,
                "total_pallets": getattr(row, 'total_pallets', 0.0)
            }
            for row in results
        ]
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.post("/api/registry/reset")
async def reset_registry(db: Session = Depends(get_db)):
    """
    Resets the database by dropping and recreating tables, and clears processed JSON files.
    """
    try:
        # 1. Clear database tables
        Base.metadata.drop_all(bind=engine)
        init_db()
        
        # 2. Clear processed JSON folder
        processed_dir = sys_config.temp_uploads_dir / "processed"
        if processed_dir.exists():
            for f in processed_dir.glob("*.json"):
                try:
                    f.unlink()
                except:
                    pass
                    
        return {"status": "success", "message": "Database and processed files have been reset."}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to reset registry: {str(e)}"})

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
    bundle_dir_name: str = ""

@app.post("/api/template/analyze")
def analyze_template(file: UploadFile = File(...)):
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = TEMP_DIR / file.filename
    
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        analysis_output_path = TEMP_DIR / f"{file.filename}_analysis.json"
        
        # Use Orchestrator
        json_output = orchestrator.analyze_template(temp_path, legacy_format=True)
        analysis_data = json.loads(json_output)
        
        # Save to analysis_output_path for get_missing_headers to read
        with open(analysis_output_path, 'w', encoding='utf-8') as f:
            f.write(json_output)
            
        missing_headers = get_missing_headers(str(analysis_output_path))
        
        # Clean up analysis file immediately, keep excel for next step
        if analysis_output_path.exists():
            analysis_output_path.unlink()
            
        return {
            "missing_headers": missing_headers,
            "warnings": analysis_data.get("warnings", []),
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
def generate_template(config: TemplateConfig):
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
            custom_prefix=config.file_prefix,
            bundle_dir_name=config.bundle_dir_name or None
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
    List all available template variants within bundles.
    """
    bundled_dir = sys_config.bundled_dir
    templates = []
    
    if bundled_dir.exists():
        for bundle_folder in bundled_dir.iterdir():
            if bundle_folder.is_dir():
                # Find all *_template.json in the bundle folder
                for template_json_path in bundle_folder.glob("*_template.json"):
                    try:
                        stats = template_json_path.stat()
                        source_file = "Unknown"
                        try:
                            with open(template_json_path, 'r', encoding='utf-8') as f:
                                data = json.load(f)
                                source_file = data.get("fingerprint", {}).get("source_file", "Unknown")
                        except Exception:
                            pass

                        # Extract name, e.g., "MT_KH" from "MT_KH_template.json"
                        name = template_json_path.name.replace("_template.json", "")

                        templates.append({
                            "name": name,
                            "bundle_name": bundle_folder.name,
                            "path": str(template_json_path),
                            "modified": datetime.datetime.fromtimestamp(stats.st_mtime).isoformat(),
                            "source_file": source_file
                        })
                    except Exception:
                        pass
                        
    return templates

@app.get("/api/template/view")
async def view_template(name: str, bundle: Optional[str] = None):
    """
    Get the content of a specific template JSON.
    """
    bundled_dir = sys_config.bundled_dir
    # Sanitize name to prevent traversal
    safe_name = Path(name).name
    
    if bundle:
        safe_bundle = Path(bundle).name
        template_dir = bundled_dir / safe_bundle
    else:
        # Fallback to older logic where bundle == name
        safe_name = Path(name).name
        template_dir = bundled_dir / safe_name
        
        # Fallback 2: Search for the bundle containing this template
        if not template_dir.exists() or not template_dir.is_dir():
            if bundled_dir.exists() and bundled_dir.is_dir():
                for b_dir in bundled_dir.iterdir():
                    if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                        template_dir = b_dir
                        break
                        
    template_path = template_dir / f"{safe_name}_template.json"
    
    if not template_path.exists():
        return JSONResponse(status_code=404, content={"error": "Template JSON not found"})
        
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to read template: {str(e)}"})


class CellOverrideRequest(BaseModel):
    """Request body for saving a mode-dependent cell override."""
    template_name: str
    bundle_name: str = ""
    sheet_name: str
    cell_address: str   # e.g. "A1"
    mode: str           # "daf", "standard", etc.
    value: str          # new value for this mode


@app.patch("/api/template/cell")
async def update_template_cell(req: CellOverrideRequest):
    """
    Save a mode-dependent override for a header cell.
    
    Converts the cell value from a plain string to a dict like:
        {"default": "INVOICE", "daf": "DAF"}
    If the cell is already a dict, just updates the specified mode key.
    """
    bundled_dir = sys_config.bundled_dir
    safe_name = Path(req.template_name).name

    # Resolve template directory (same logic as view_template)
    if req.bundle_name:
        template_dir = bundled_dir / Path(req.bundle_name).name
    else:
        template_dir = bundled_dir / safe_name
        if not template_dir.exists() or not template_dir.is_dir():
            if bundled_dir.exists() and bundled_dir.is_dir():
                for b_dir in bundled_dir.iterdir():
                    if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                        template_dir = b_dir
                        break

    template_path = template_dir / f"{safe_name}_template.json"
    if not template_path.exists():
        return JSONResponse(status_code=404, content={"error": "Template JSON not found"})

    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to read template: {str(e)}"})

    # Validate sheet exists (sheets are nested under template_layout)
    template_layout = data.get("template_layout", {})
    if req.sheet_name not in template_layout:
        return JSONResponse(status_code=404, content={"error": f"Sheet '{req.sheet_name}' not found in template"})

    sheet_data = template_layout[req.sheet_name]
    header_content = sheet_data.get("header_content", {})
    current = header_content.get(req.cell_address)

    # Build the mode-dependent value
    if isinstance(current, dict):
        # Already a mode map — update the mode key
        current[req.mode] = req.value
    elif current is not None:
        # Plain string — convert to mode map with 'default' as the original
        header_content[req.cell_address] = {"default": current, req.mode: req.value}
    else:
        # Cell doesn't exist yet — create with just the mode key
        header_content[req.cell_address] = {"default": "", req.mode: req.value}

    sheet_data["header_content"] = header_content

    # Save back
    try:
        with open(template_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return {"status": "success", "cell": req.cell_address, "value": header_content[req.cell_address]}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to save template: {str(e)}"})

@app.delete("/api/template/{name}")
async def delete_template(name: str, bundle: Optional[str] = None):
    """
    Delete a specific template bundle.
    """
    bundled_dir = sys_config.bundled_dir
    
    if bundle:
        safe_bundle = Path(bundle).name
        template_dir = bundled_dir / safe_bundle
    else:
        # Fallback
        safe_name = Path(name).name
        template_dir = bundled_dir / safe_name
        
        # Fallback 2: Search for the bundle containing this template
        if not template_dir.exists() or not template_dir.is_dir():
            if bundled_dir.exists() and bundled_dir.is_dir():
                for b_dir in bundled_dir.iterdir():
                    if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                        template_dir = b_dir
                        break
    
    if not template_dir.exists() or not template_dir.is_dir():
        return JSONResponse(status_code=404, content={"error": "Template bundle not found"})
        
    try:
        shutil.rmtree(template_dir)
        return {"status": "success", "message": "Template bundle deleted successfully"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to delete template bundle: {str(e)}"})

# --- Log Viewer API ---
from core.logger_config import clear_session_log

@app.get("/api/logs/current")
async def get_current_log():
    """
    Read and return the contents of current_session.log.

    Returns:
        JSON with 'content' (log text) and 'lines' (line count).
    """
    log_file = sys_config.run_log_dir / "current_session.log"

    if not log_file.exists():
        return {"content": "", "lines": 0}

    try:
        with open(log_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        line_count = content.count('\n')
        return {"content": content, "lines": line_count}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to read log: {str(e)}"})


@app.post("/api/logs/clear")
async def clear_current_log():
    """
    Clear the current session log file.

    Uses the existing clear_session_log() utility from logger_config.
    """
    try:
        clear_session_log()
        return {"status": "ok", "message": "Session log cleared."}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to clear log: {str(e)}"})
