import logging
import json
import shutil
import datetime
from fastapi import APIRouter, UploadFile, File
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import List, Optional, Dict
from pathlib import Path
from core.system_config import sys_config
from core.orchestrator import Orchestrator
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, range_boundaries

router = APIRouter(prefix="/api", tags=["templates"])
logger = logging.getLogger(__name__)
orchestrator = Orchestrator()

class TemplateConfig(BaseModel):
    file_prefix: str
    user_mappings: dict
    temp_filename: str
    bundle_dir_name: str = ""
    confirmed_footers: List[str] = []
    pricing_mode: str = "standard"  # 'standard' or 'net'

class CellOverrideRequest(BaseModel):
    template_name: str
    bundle_name: str = ""
    sheet_name: str
    cell_address: str
    overrides: Dict[str, str]

class TemplateNotesRequest(BaseModel):
    template_name: str
    bundle_name: str = ""
    notes: str

# --- Helpers ---

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
        for sheet in analysis_data.get('sheets', []):
            for header_pos in sheet.get('header_positions', []):
                col_id = header_pos.get('col_id', '')
                header_text = header_pos.get('keyword', '')
                if col_id.startswith("col_unknown"):
                    missing_headers.append({"text": header_text, "suggestion": get_header_suggestions(header_text)})
        return missing_headers
    except Exception:
        logger.exception("Failed to read missing headers from %s", analysis_file_path)
        return []

def get_missing_footers(analysis_file_path: str):
    try:
        with open(analysis_file_path, 'r', encoding='utf-8') as f:
            analysis_data = json.load(f)
        missing_footers = []
        for sheet in analysis_data.get('sheets', []):
            uf = sheet.get("unconfirmed_footer")
            if uf: missing_footers.append(uf)
        return list(set(missing_footers))
    except Exception:
        logger.exception("Failed to read missing footers from %s", analysis_file_path)
        return []

def update_mapping_config(new_mappings: dict):
    try:
        mapping_path = sys_config.mapping_config_path
        mapping_data = {"header_text_mappings": {"mappings": {}}}
        if mapping_path.exists():
            with open(mapping_path, 'r', encoding='utf-8') as f:
                mapping_data = json.load(f)
        if "header_text_mappings" not in mapping_data: mapping_data["header_text_mappings"] = {"mappings": {}}
        filtered = {k: v for k, v in new_mappings.items() if v and v != "col_unknown"}
        if filtered:
            mapping_data["header_text_mappings"]["mappings"].update(filtered)
            with open(mapping_path, 'w', encoding='utf-8') as f:
                json.dump(mapping_data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        logger.exception("Failed to update mapping config")
        return False

def read_table_info_from_config(template_dir: Path) -> dict:
    result = {}
    config_files = list(template_dir.glob("*_config.json"))
    if not config_files: return result
    try:
        with open(config_files[0], 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        layout = cfg.get("layout_bundle", {})
        fallback = layout.get("defaults", {}).get("data_flow", {}).get("mappings", {}).get("col_desc", {}).get("fallback", {})
        if fallback: result["fallback_description"] = fallback
        for _, sheet_data in layout.items():
            if not isinstance(sheet_data, dict): continue
            bf = sheet_data.get("footer", {}).get("add_ons", {}).get("before_footer", {})
            if bf.get("enabled") and bf.get("text"):
                result["hs_code"] = bf["text"]
                break
    except Exception:
        logger.exception("Failed to read table info from config in %s", template_dir)
    return result

# --- Routes ---

@router.post("/template/analyze")
def analyze_template(file: UploadFile = File(...)):
    temp_dir = sys_config.temp_uploads_dir
    temp_path = temp_dir / file.filename
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        json_output = orchestrator.analyze_template(temp_path, legacy_format=True)
        analysis_path = temp_dir / f"{file.filename}_analysis.json"
        with open(analysis_path, 'w', encoding='utf-8') as f: f.write(json_output)
        res = {
            "missing_headers": get_missing_headers(str(analysis_path)),
            "missing_footers": get_missing_footers(str(analysis_path)),
            "warnings": json.loads(json_output).get("warnings", []),
            "temp_filename": file.filename,
            "suggested_prefix": file.filename.split('.')[0]
        }
        if analysis_path.exists(): analysis_path.unlink()
        return res
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/template/generate")
def generate_template(config: TemplateConfig):
    try:
        if config.user_mappings and not update_mapping_config(config.user_mappings):
            return JSONResponse(status_code=500, content={"error": "Mapping update failed"})
        
        # Footer label update
        if config.confirmed_footers:
            mapping_path = sys_config.mapping_config_path
            if mapping_path.exists():
                with open(mapping_path, 'r', encoding='utf-8') as f: data = json.load(f)
                if "footer_label_mappings" not in data: data["footer_label_mappings"] = {"keywords": []}
                existing = data["footer_label_mappings"].get("keywords", [])
                for fm in config.confirmed_footers:
                    if fm not in existing: existing.append(fm)
                data["footer_label_mappings"]["keywords"] = existing
                with open(mapping_path, 'w', encoding='utf-8') as f: json.dump(data, f, indent=4, ensure_ascii=False)

        temp_path = sys_config.temp_uploads_dir / config.temp_filename
        result_path = orchestrator.generate_blueprint_bundle(
            template_path=temp_path,
            output_dir=sys_config.bundled_dir,
            custom_prefix=config.file_prefix,
            runtime_mappings=config.user_mappings,
            bundle_dir_name=config.bundle_dir_name or None,
            pricing_mode=config.pricing_mode
        )
        
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception as cleanup_err:
            logger.warning(f"Failed to delete temporary blueprint file {temp_path}: {cleanup_err}")
            
        return {"status": "success", "bundle_path": str(result_path.parent)}
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.get("/templates")
async def list_templates():
    bundled_dir = sys_config.bundled_dir
    templates = []
    if bundled_dir.exists():
        for b_folder in bundled_dir.iterdir():
            if b_folder.is_dir():
                for t_json in b_folder.glob("*_template.json"):
                    try:
                        stats = t_json.stat()
                        with open(t_json, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            source = data.get("fingerprint", {}).get("source_file", "Unknown")
                        templates.append({
                            "name": t_json.name.replace("_template.json", ""),
                            "bundle_name": b_folder.name,
                            "modified": datetime.datetime.fromtimestamp(stats.st_mtime).isoformat(),
                            "source_file": source
                        })
                    except Exception:
                        logger.exception("Failed to read template %s", t_json)
    return templates

@router.get("/template/view")
async def view_template(name: str, bundle: Optional[str] = None):
    bundled_dir = sys_config.bundled_dir
    safe_name = Path(name).name
    template_dir = bundled_dir / (Path(bundle).name if bundle else safe_name)
    if not template_dir.exists():
        for b_dir in bundled_dir.iterdir():
            if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                template_dir = b_dir; break
    
    t_path = template_dir / f"{safe_name}_template.json"
    if not t_path.exists(): return JSONResponse(status_code=404, content={"error": "Not found"})
    try:
        with open(t_path, 'r', encoding='utf-8') as f: data = json.load(f)
        info = read_table_info_from_config(template_dir)
        if info:
            if "table_info" not in data: data["table_info"] = {}
            data["table_info"].update(info)
        return data
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.patch("/api/template/cell")
# Note: I used /api/template/cell in main.py but the router prefix is /api, so it should be /template/cell
@router.patch("/template/cell")
async def update_template_cell(req: CellOverrideRequest):
    bundled_dir = sys_config.bundled_dir
    safe_name = Path(req.template_name).name
    template_dir = bundled_dir / (Path(req.bundle_name).name if req.bundle_name else safe_name)
    if not template_dir.exists():
        for b_dir in bundled_dir.iterdir():
            if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                template_dir = b_dir; break
    
    t_path = template_dir / f"{safe_name}_template.json"
    if not t_path.exists(): return JSONResponse(status_code=404, content={"error": "Not found"})
    try:
        with open(t_path, 'r', encoding='utf-8') as f: data = json.load(f)
        sheet = data.get("template_layout", {}).get(req.sheet_name)
        if not sheet: return JSONResponse(status_code=404, content={"error": "Sheet not found"})
        
        def get_max_row(content, merges):
            max_r = 0
            for addr in content.keys():
                _, r = coordinate_from_string(addr); max_r = max(max_r, r)
            m_list = merges if isinstance(merges, list) else merges.keys() if isinstance(merges, dict) else []
            for m in m_list:
                _, _, _, mr = range_boundaries(m); max_r = max(max_r, mr)
            return max_r

        h_content = sheet.get("header_content", {})
        h_max = get_max_row(h_content, sheet.get("header_merges", []))
        col_letter, row_val = coordinate_from_string(req.cell_address)
        col_idx = column_index_from_string(col_letter)
        is_f = row_val > h_max

        if is_f:
            rel = row_val - h_max - 1
            f_rows = sheet.get("footer_rows", [])
            row = next((r for r in f_rows if r.get('relative_index') == rel), None)
            if not row:
                row = {"relative_index": rel, "cells": [], "merges": []}
                f_rows.append(row)
                sheet["footer_rows"] = sorted(f_rows, key=lambda x: x.get('relative_index', 0))
            cells = row.get("cells", [])
            cell = next((c for c in cells if c.get('col_index') == col_idx), None)
            if not cell:
                cell = {"col_index": col_idx, "value": ""}
                cells.append(cell)
                row["cells"] = sorted(cells, key=lambda x: x.get('col_index', 1))
            val = cell.get("value")
            curr_map = val if isinstance(val, dict) else {"default": str(val) if val is not None else ""}
            for m, v in req.overrides.items():
                if v is None or (isinstance(v, str) and not v.strip()):
                    if m in curr_map: del curr_map[m]
                else: curr_map[m] = v
            if len(curr_map) == 1 and "default" in curr_map: cell["value"] = curr_map["default"]
            elif not curr_map: cell["value"] = ""
            else: cell["value"] = curr_map
        else:
            val = h_content.get(req.cell_address)
            curr_map = val if isinstance(val, dict) else {"default": str(val) if val is not None else ""}
            for m, v in req.overrides.items():
                if v is None or (isinstance(v, str) and not v.strip()):
                    if m in curr_map: del curr_map[m]
                else: curr_map[m] = v
            if len(curr_map) == 1 and "default" in curr_map: h_content[req.cell_address] = curr_map["default"]
            elif not curr_map:
                if req.cell_address in h_content: del h_content[req.cell_address]
            else: h_content[req.cell_address] = curr_map

        with open(t_path, 'w', encoding='utf-8') as f: json.dump(data, f, indent=2, ensure_ascii=False)
        return {"status": "success"}
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.patch("/template/notes")
async def update_template_notes(req: TemplateNotesRequest):
    bundled_dir = sys_config.bundled_dir
    safe_name = Path(req.template_name).name
    template_dir = bundled_dir / (Path(req.bundle_name).name if req.bundle_name else safe_name)
    if not template_dir.exists():
        for b_dir in bundled_dir.iterdir():
            if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                template_dir = b_dir; break
    t_path = template_dir / f"{safe_name}_template.json"
    if not t_path.exists(): return JSONResponse(status_code=404, content={"error": "Not found"})
    try:
        with open(t_path, 'r', encoding='utf-8') as f: data = json.load(f)
        data["notes"] = req.notes
        with open(t_path, 'w', encoding='utf-8') as f: json.dump(data, f, indent=2, ensure_ascii=False)
        return {"status": "success"}
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})

@router.delete("/template/{name}")
async def delete_template(name: str, bundle: Optional[str] = None):
    bundled_dir = sys_config.bundled_dir
    safe_name = Path(name).name
    template_dir = bundled_dir / (Path(bundle).name if bundle else safe_name)
    
    # Try to find the directory if the direct path doesn't exist
    if not template_dir.exists():
        for b_dir in bundled_dir.iterdir():
            if b_dir.is_dir() and (b_dir / f"{safe_name}_template.json").exists():
                template_dir = b_dir; break
    
    if not template_dir.exists(): return JSONResponse(status_code=404, content={"error": "Not found"})
    
    try:
        # Surgical deletion of specific files
        files_to_delete = [
            template_dir / f"{safe_name}_template.json",
            template_dir / f"{safe_name}_config.json",
            template_dir / f"{safe_name}.xlsx"
        ]
        
        deleted_count = 0
        for f in files_to_delete:
            if f.exists():
                f.unlink()
                deleted_count += 1
        
        # If the directory is now empty, remove it too
        if template_dir.exists() and not any(template_dir.iterdir()):
            shutil.rmtree(template_dir)
            
        if deleted_count == 0:
             return JSONResponse(status_code=404, content={"error": f"No files found for template '{safe_name}' in bundle '{template_dir.name}'"})

        return {"status": "success", "deleted_files": deleted_count}
    except Exception as e: return JSONResponse(status_code=500, content={"error": str(e)})
