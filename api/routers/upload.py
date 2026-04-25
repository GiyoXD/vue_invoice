import logging
import io
from fastapi import APIRouter, UploadFile, File
from fastapi.responses import JSONResponse
from pathlib import Path
from core.system_config import sys_config
from core.orchestrator import Orchestrator
from core.data_parser.data_processor import DataValidationError
import json

router = APIRouter(prefix="/api", tags=["upload"])
logger = logging.getLogger(__name__)
orchestrator = Orchestrator()

@router.post("/upload")
def upload_excel(file: UploadFile = File(...)):
    """
    Uploads an Excel file and processes it to JSON.
    Returns the identifier, json path, and asset availability status.
    """
    try:
        print(f"DEBUG: Received upload request for {file.filename}")
        
        # Read the file into memory
        file_bytes = file.file.read()
        buffer = io.BytesIO(file_bytes)
            
        # Process to JSON using Orchestrator
        upload_dir = sys_config.temp_uploads_dir
        json_output_dir = upload_dir / "processed"
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
            asset_status["message"] = "Ready to generate invoice."
            
            # Read pricing_mode from config for frontend
            try:
                with open(assets.config_path, 'r', encoding='utf-8') as cf:
                    config_data = json.load(cf)
                asset_status["pricing_mode"] = config_data.get("_meta", {}).get("pricing_mode", "standard")
            except Exception as pm_err:
                logger.warning(f"Could not read pricing_mode from config: {pm_err}")
                asset_status["pricing_mode"] = "standard"
        else:
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
