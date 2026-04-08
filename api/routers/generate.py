import logging
import json
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import List, Optional, Any
from pathlib import Path
from core.system_config import sys_config
from core.orchestrator import Orchestrator

router = APIRouter(prefix="/api", tags=["generate"])
logger = logging.getLogger(__name__)
orchestrator = Orchestrator()

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

@router.post("/generate")
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

        # Define base output path
        output_dir = sys_config.output_dir
        base_output_dir = output_dir / request.identifier

        # Default Template/Config dirs
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

        # Price adjustments
        from core.invoice_generator.utils.aggregation_modifier import apply_aggregation_adjustment
        price_adj = request.price_adjustment
        if price_adj:
            full_data = apply_aggregation_adjustment(full_data, price_adj)
        
        # PERSIST changes to disk
        try:
            with open(json_path_obj, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, indent=4)
        except Exception as e:
            logger.warning(f"Failed to persist updated JSON to {json_path_obj}: {e}")
        
        results = []
        errors = []
        generated_files = []
        processed_any = False

        # Define mode tasks
        mode_tasks = []
        if request.generate_standard:
            mode_tasks.append({"suffix": "", "flags": [], "name": "Standard Invoice"})
        if request.generate_custom:
            mode_tasks.append({"suffix": "_Custom", "flags": ["--custom"], "name": "Custom Invoice"})
        if request.generate_daf:
            mode_tasks.append({"suffix": "_DAF", "flags": ["--DAF"], "name": "DAF Invoice"})

        if not mode_tasks:
             return JSONResponse(status_code=400, content={"error": "No invoice type selected."})

        # Determine variant tasks
        variant_tasks = []
        if request.generate_kh or request.generate_vn:
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
        
        if not variant_tasks:
            variant_tasks = [{"suffix": "", "config_path": None, "template_path": None}]

        # Final loop
        for variant in variant_tasks:
            variant_suffix = variant["suffix"]
            for task in mode_tasks:
                try:
                    filename = f"{request.identifier}_Invoice{variant_suffix}{task['suffix']}.xlsx"
                    output_path = base_output_dir / filename

                    flags = list(task['flags'])
                    if variant.get("config_path"):
                        flags.extend(["--config", str(variant["config_path"])])
                    if variant.get("template_path"):
                        flags.extend(["--template", str(variant["template_path"])])

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
                        fname, fbytes = result
                        results.append(fname)
                        generated_files.append((fname, fbytes))
                        processed_any = True
                except Exception as e:
                    task_name = f"{variant_suffix.lstrip('_')} {task['name']}" if variant_suffix else task['name']
                    errors.append(f"Failed to generate {task_name}: {str(e)}")

        if not processed_any and errors:
             status_code = 500
             error_message = "All generation tasks failed."
             config_errors = [err for err in errors if "CRITICAL: No 'header_row'" in err]
             if config_errors:
                 status_code = 422
                 error_message = config_errors[0]

             return JSONResponse(status_code=status_code, content={
                "error": error_message, 
                "details": errors
            })

        final_payload_files = []
        if generated_files:
            import zipfile
            import base64
            import io
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
