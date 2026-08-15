import logging
import io
import os
import json
import datetime
from pathlib import Path
from fastapi import APIRouter, UploadFile, File, Form, Depends
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from sqlalchemy.orm import Session

from core.system_config import sys_config
from core.orchestrator import Orchestrator
from core.data_parser.data_processor import DataValidationError
from core.database.db_manager import get_db
from core.database.models.system import SystemSetting

router = APIRouter(prefix="/api", tags=["upload"])
logger = logging.getLogger(__name__)
orchestrator = Orchestrator()


class FolderRequest(BaseModel):
    folder_path: str


class OpenFileRequest(BaseModel):
    filename: str


class ProcessExistingRequest(BaseModel):
    filename: str
    ignore_tare: bool = False
    ignore_cbm: bool = False


def get_source_folder(db: Session) -> Path:
    try:
        setting = db.query(SystemSetting).filter(SystemSetting.key == "source_files_folder").first()
        if setting and setting.value_json:
            folder_str = json.loads(setting.value_json)
            folder_path = Path(folder_str)
            if folder_path.exists() and folder_path.is_dir():
                return folder_path
    except Exception as e:
        logger.warning(f"Failed to load source_files_folder setting: {e}")
    
    fallback = sys_config.temp_uploads_dir
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def set_source_folder(db: Session, folder_str: str) -> Path:
    if not folder_str or not folder_str.strip():
        raise ValueError("Folder path cannot be empty")
    folder_path = Path(folder_str.strip()).resolve()
    if not folder_path.exists() or not folder_path.is_dir():
        raise ValueError(f"Directory does not exist: {folder_str.strip()}")
    
    setting = db.query(SystemSetting).filter(SystemSetting.key == "source_files_folder").first()
    if not setting:
        setting = SystemSetting(key="source_files_folder", value_json=json.dumps(str(folder_path)))
        db.add(setting)
    else:
        setting.value_json = json.dumps(str(folder_path))
    db.commit()
    return folder_path


def _build_upload_response(filename: str, json_path: Path, identifier: str) -> dict:
    default_inv_no = Path(filename).stem
    
    # === CHECK ASSET AVAILABILITY ===
    from core.invoice_generator.resolvers import InvoiceAssetResolver
    
    resolver = InvoiceAssetResolver(
        base_config_dir=sys_config.registry_dir,
        base_template_dir=sys_config.templates_dir
    )
    
    assets = resolver.resolve_assets_for_input_file(str(json_path))
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
        asset_status["blueprint_name"] = assets.config_data.get("_meta", {}).get("description") or getattr(assets, "name", None) or identifier
        asset_status["message"] = "Ready to generate invoice."
        
        # Read pricing_mode from config for frontend
        try:
            asset_status["pricing_mode"] = assets.config_data.get("_meta", {}).get("pricing_mode", "standard") if assets.config_data else "standard"
        except Exception as pm_err:
            logger.warning(f"Could not read pricing_mode from config: {pm_err}")
            asset_status["pricing_mode"] = "standard"
    else:
        asset_status["message"] = f"No configuration found in database for '{identifier}'."
    
    # --- Read warnings and parsed data from generated JSON ---
    warnings_list = []
    parsed_json_data = None
    try:
        if json_path and Path(json_path).exists():
            with open(json_path, 'r', encoding='utf-8') as f:
                parsed_json_data = json.load(f)
                if isinstance(parsed_json_data, dict):
                    metadata = parsed_json_data.get('metadata')
                    if isinstance(metadata, dict):
                        raw_warnings = metadata.get('warnings')
                        if isinstance(raw_warnings, list):
                            warnings_list = raw_warnings
    except Exception as e:
        logger.warning(f"Could not read warnings/data from JSON output: {e}")
        warnings_list = []

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
        "file_name": filename,
        "identifier": identifier,
        "json_path": str(json_path),
        "default_inv_no": default_inv_no,
        "warnings": warnings_list,
        "asset_status": asset_status,
        "parsed_data": parsed_json_data,
        "message": "File processed successfully"
    }


@router.get("/source-folder")
def get_source_folder_endpoint(db: Session = Depends(get_db)):
    folder = get_source_folder(db)
    return {"folder_path": str(folder)}


@router.post("/source-folder")
def set_source_folder_endpoint(req: FolderRequest, db: Session = Depends(get_db)):
    raw_path = (req.folder_path or "").strip()
    if not raw_path:
        return JSONResponse(
            status_code=400,
            content={"error": "Folder path cannot be empty", "step": "Source Folder"}
        )
    target_path = Path(raw_path).resolve()
    if not target_path.exists() or not target_path.is_dir():
        return JSONResponse(
            status_code=400,
            content={"error": f"Directory does not exist: {raw_path}", "step": "Source Folder"}
        )
    try:
        folder = set_source_folder(db, str(target_path))
        return {"status": "success", "folder_path": str(folder)}
    except Exception as e:
        return JSONResponse(
            status_code=400,
            content={"error": str(e), "step": "Source Folder"}
        )


@router.get("/source-files")
def list_source_files(db: Session = Depends(get_db)):
    folder = get_source_folder(db)
    folder.mkdir(parents=True, exist_ok=True)
    files = []
    try:
        for entry in folder.iterdir():
            try:
                if entry.is_file() and entry.suffix.lower() in [".xlsx", ".xls"]:
                    if entry.name.startswith("~$"):
                        continue
                    stat = entry.stat()
                    files.append({
                        "filename": entry.name,
                        "size_bytes": stat.st_size,
                        "modified_at": datetime.datetime.fromtimestamp(stat.st_mtime).isoformat(),
                        "modified_timestamp": stat.st_mtime
                    })
            except (FileNotFoundError, PermissionError, OSError) as fe:
                logger.warning(f"Skipping inaccessible file {entry.name}: {fe}")
                continue
        # Sort newest first
        files.sort(key=lambda x: x["modified_timestamp"], reverse=True)
    except Exception as e:
        logger.error(f"Error listing source files: {e}")
        return JSONResponse(
            status_code=500,
            content={"error": f"Failed to list files: {str(e)}", "step": "Source Files"}
        )

    return {
        "folder_path": str(folder),
        "files": files
    }


@router.post("/open-file")
def open_file_in_excel(req: OpenFileRequest, db: Session = Depends(get_db)):
    file_name = Path(req.filename).name
    ext = Path(file_name).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        return JSONResponse(
            status_code=400,
            content={"error": "Only .xlsx and .xls files are supported", "step": "Open File"}
        )
    folder = get_source_folder(db)
    file_path = folder / file_name
    if not file_path.exists() or not file_path.is_file():
        return JSONResponse(
            status_code=404,
            content={"error": f"File not found: {file_name}", "step": "Open File"}
        )
    
    try:
        if hasattr(os, "startfile"):
            os.startfile(str(file_path))
        else:
            import subprocess
            import sys
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(file_path)])
            else:
                subprocess.Popen(["xdg-open", str(file_path)])
        return {"status": "success", "message": f"Opened {file_name}"}
    except Exception as e:
        logger.error(f"Error opening file {file_path}: {e}")
        return JSONResponse(
            status_code=500,
            content={"error": f"Failed to open file: {str(e)}", "step": "Open File"}
        )


@router.post("/process-existing")
def process_existing_file(req: ProcessExistingRequest, db: Session = Depends(get_db)):
    file_name = Path(req.filename).name
    ext = Path(file_name).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        return JSONResponse(
            status_code=400,
            content={"error": "Only .xlsx and .xls files are supported", "step": "Process Existing File"}
        )
    folder = get_source_folder(db)
    file_path = folder / file_name
    if not file_path.exists() or not file_path.is_file():
        return JSONResponse(
            status_code=404,
            content={"error": f"File not found: {file_name}", "step": "Process Existing File"}
        )

    try:
        upload_dir = sys_config.temp_uploads_dir
        json_output_dir = upload_dir / "processed"
        json_output_dir.mkdir(parents=True, exist_ok=True)

        json_path, identifier = orchestrator.process_excel_to_json(
            file_path,
            json_output_dir,
            input_filename_override=file_name,
            ignore_tare_warning=req.ignore_tare,
            ignore_cbm_warning=req.ignore_cbm
        )
        return _build_upload_response(file_name, json_path, identifier)
    except DataValidationError as ve:
        return JSONResponse(status_code=422, content={
            "error": str(ve),
            "step": "Data Validation"
        })
    except Exception as e:
        logger.error("Processing existing file failed", exc_info=True)
        return JSONResponse(status_code=500, content={
            "error": f"Internal server error during processing: {str(e)}",
            "step": "Process Existing File"
        })


@router.post("/upload")
def upload_excel(
    file: UploadFile = File(...),
    ignore_tare: bool = Form(False),
    ignore_cbm: bool = Form(False),
    db: Session = Depends(get_db)
):
    """
    Uploads an Excel file and processes it to JSON.
    Returns the identifier, json path, and asset availability status.
    """
    try:
        raw_name = file.filename or "upload.xlsx"
        filename = Path(raw_name).name
        ext = Path(filename).suffix.lower()
        if ext not in [".xlsx", ".xls"]:
            return JSONResponse(
                status_code=400,
                content={"error": "Only .xlsx and .xls files are supported", "step": "File Upload"}
            )
        logger.debug(f"Received upload request for {filename}")
        
        # Read the file into memory
        MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
        file_bytes = file.file.read(MAX_FILE_SIZE + 1)
        if len(file_bytes) > MAX_FILE_SIZE:
            file.file.close()
            return JSONResponse(
                status_code=413,
                content={"error": "File size exceeds maximum limit of 50MB", "step": "File Upload"}
            )
        if len(file_bytes) == 0:
            return JSONResponse(
                status_code=400,
                content={"error": "Uploaded file is empty", "step": "File Upload"}
            )
            
        # Save uploaded file to source folder
        source_folder = get_source_folder(db)
        source_folder.mkdir(parents=True, exist_ok=True)
        saved_file_path = source_folder / filename
        with open(saved_file_path, "wb") as f:
            f.write(file_bytes)

        buffer = io.BytesIO(file_bytes)
            
        # Process to JSON using Orchestrator
        upload_dir = sys_config.temp_uploads_dir
        json_output_dir = upload_dir / "processed"
        json_output_dir.mkdir(parents=True, exist_ok=True)

        json_path, identifier = orchestrator.process_excel_to_json(
            buffer, 
            json_output_dir,
            input_filename_override=filename,
            ignore_tare_warning=ignore_tare,
            ignore_cbm_warning=ignore_cbm
        )
        
        return _build_upload_response(filename, json_path, identifier)

    except DataValidationError as ve:
        return JSONResponse(status_code=422, content={
            "error": str(ve),
            "step": "Data Validation"
        })
    except Exception as e:
        logger.error("Upload failed", exc_info=True)
        return JSONResponse(status_code=500, content={
            "error": f"Internal server error during upload: {str(e)}",
            "step": "Upload & Parse"
        })
