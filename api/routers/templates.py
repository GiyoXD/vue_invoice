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
from core.invoice_generator.extractors.template_client_profile_parser import TemplateClientProfileParser
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, range_boundaries

router = APIRouter(prefix="/api", tags=["templates"])
logger = logging.getLogger(__name__)
orchestrator = Orchestrator()

class TemplateConfig(BaseModel):
    customer_code: str
    locale: str = "KH"
    user_mappings: dict
    temp_filename: str
    bundle_dir_name: str = ""
    confirmed_footers: List[str] = []
    pricing_mode: str = "standard"  # 'standard' or 'net'
    ignore_missing_description: bool = False

class CellOverrideRequest(BaseModel):
    customer_code: str
    locale: str = "KH"
    sheet_name: str
    cell_address: str
    overrides: Dict[str, str]

class TemplateNotesRequest(BaseModel):
    customer_code: str
    locale: str = "KH"
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
        from core.database.db_manager import get_global_mapping_config, save_global_mapping_config
        mapping_data = get_global_mapping_config()
        if "header_text_mappings" not in mapping_data: mapping_data["header_text_mappings"] = {"mappings": {}}
        filtered = {k: v for k, v in new_mappings.items() if v and v != "col_unknown"}
        if filtered:
            mapping_data["header_text_mappings"]["mappings"].update(filtered)
            save_global_mapping_config(mapping_data)
        return True
    except Exception:
        logger.exception("Failed to update mapping config")
        return False


# --- Routes ---

@router.post("/template/analyze")
def analyze_template(file: UploadFile = File(...), ignore_missing_description: bool = False):
    temp_dir = sys_config.temp_uploads_dir
    temp_path = temp_dir / file.filename
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        json_output = orchestrator.analyze_template(temp_path, legacy_format=True, ignore_missing_description=ignore_missing_description)
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

def save_blueprint_to_db(customer_code: str, locale: str, config_data: dict, template_json_data: dict, xlsx_bytes: bytes, filename: str):
    """
    Saves the generated blueprint configuration and template directly to the SQLite database.
    """
    from core.database.db_manager import SessionLocal, Blueprint, BlueprintTemplate
    
    config_str = json.dumps(config_data, ensure_ascii=False)
    template_str = json.dumps(template_json_data, ensure_ascii=False)
    description = config_data.get("_meta", {}).get("description", f"Generated blueprint for {customer_code}_{locale}")
    
    db = SessionLocal()
    try:
        existing = db.query(Blueprint).filter(
            Blueprint.customer_code == customer_code,
            Blueprint.locale == locale
        ).first()
        
        if existing:
            existing.description = description
            existing.config_json = config_str
            existing.template_json = template_str
            if existing.template_binary:
                existing.template_binary.filename = filename
                existing.template_binary.xlsx_blob = xlsx_bytes
            else:
                existing.template_binary = BlueprintTemplate(
                    filename=filename,
                    xlsx_blob=xlsx_bytes
                )
        else:
            blueprint = Blueprint(
                customer_code=customer_code,
                locale=locale,
                description=description,
                config_json=config_str,
                template_json=template_str
            )
            blueprint.template_binary = BlueprintTemplate(
                filename=filename,
                xlsx_blob=xlsx_bytes
            )
            db.add(blueprint)
            
        db.commit()
        
        # Clear temp cache for this customer
        try:
            from core.system_config import sys_config
            temp_dir = sys_config.temp_uploads_dir / "runtime_blueprints" / f"{customer_code}_{locale}"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
        except Exception:
            pass
            
    except Exception as e:
        db.rollback()
        raise e
    finally:
        db.close()

@router.post("/template/generate")
def generate_template(config: TemplateConfig):
    temp_path = sys_config.temp_uploads_dir / config.temp_filename
    try:
        if config.user_mappings and not update_mapping_config(config.user_mappings):
            return JSONResponse(status_code=500, content={"error": "Mapping update failed"})
        
        # Footer label update
        if config.confirmed_footers:
            from core.database.db_manager import get_global_mapping_config, save_global_mapping_config
            data = get_global_mapping_config()
            if "footer_label_mappings" not in data: data["footer_label_mappings"] = {"keywords": []}
            existing = data["footer_label_mappings"].get("keywords", [])
            for fm in config.confirmed_footers:
                if fm not in existing: existing.append(fm)
            data["footer_label_mappings"]["keywords"] = existing
            save_global_mapping_config(data)

        customer_code = config.customer_code
        locale = config.locale
        
        from core.database.db_manager import SessionLocal, Blueprint
        db = SessionLocal()
        existing_template_json = None
        try:
            existing = db.query(Blueprint).filter(
                Blueprint.customer_code == customer_code,
                Blueprint.locale == locale
            ).first()
            if existing:
                existing_template_json = json.loads(existing.template_json)
        finally:
            db.close()

        config_data, template_json_data, template_xlsx_bytes = orchestrator.generate_blueprint_bundle(
            template_path=temp_path,
            custom_prefix=customer_code,
            runtime_mappings=config.user_mappings,
            bundle_dir_name=config.bundle_dir_name or None,
            pricing_mode=config.pricing_mode,
            ignore_missing_description=config.ignore_missing_description,
            in_memory=True,
            existing_template_json=existing_template_json
        )
        
        # Save to SQLite and remove files from disk
        save_blueprint_to_db(
            customer_code=customer_code,
            locale=locale,
            config_data=config_data,
            template_json_data=template_json_data,
            xlsx_bytes=template_xlsx_bytes,
            filename=f"{customer_code}_{locale}.xlsx"
        )
        
        return {"status": "success", "message": "Blueprint saved to database."}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception as cleanup_err:
            logger.warning(f"Failed to delete temporary blueprint file {temp_path}: {cleanup_err}")

def _clear_cache(customer_code: str, locale: str):
    try:
        from core.system_config import sys_config
        temp_dir = sys_config.temp_uploads_dir / "runtime_blueprints" / f"{customer_code}_{locale}"
        if temp_dir.exists():
            import shutil
            shutil.rmtree(temp_dir)
            logger.info(f"Cleared cache for {customer_code}_{locale}")
    except Exception as e:
        logger.warning(f"Failed to clear cache for {customer_code}_{locale}: {e}")

@router.get("/templates")
async def list_templates():
    from core.database.db_manager import SessionLocal, Blueprint
    db = SessionLocal()
    templates = []
    try:
        rows = db.query(Blueprint).all()
        for row in rows:
            try:
                data = json.loads(row.template_json)
                source = data.get("fingerprint", {}).get("source_file", "Unknown")
                templates.append({
                    "name": f"{row.customer_code}_{row.locale}",
                    "customer_code": row.customer_code,
                    "locale": row.locale,
                    "bundle_name": row.customer_code,
                    "modified": row.updated_at.isoformat() if row.updated_at else datetime.datetime.now().isoformat(),
                    "source_file": source
                })
            except Exception:
                logger.exception("Failed to parse DB blueprint %s_%s", row.customer_code, row.locale)
    finally:
        db.close()
    return templates

@router.get("/template/view")
async def view_template(customer_code: str, locale: str = "KH"):
    from core.database.db_manager import SessionLocal, Blueprint
    
    db = SessionLocal()
    try:
        row = db.query(Blueprint).filter(
            Blueprint.customer_code == customer_code,
            Blueprint.locale == locale
        ).first()
        
        if not row:
            return JSONResponse(status_code=404, content={"error": "Not found"})
            
        data = json.loads(row.template_json)
        config_data = json.loads(row.config_json)
        info = config_data.get("table_info", {})
        if info:
            if "table_info" not in data:
                data["table_info"] = {}
            data["table_info"].update(info)

        # Inject parsed client profile so the Template Inspector can display it
        invoice_sheet = data.get("template_layout", {}).get("Invoice", {})
        header_content = invoice_sheet.get("template_header_content") or invoice_sheet.get("header_content")
        if header_content:
            parser = TemplateClientProfileParser(header_content)
            data["client_profile"] = {
                "fullname": parser.get_client_fullname(),
                "address":  parser.get_client_address(),
                "contact":  parser.get_client_contact(),
                "shipping": parser.get_shipping_method(),
            }

        return data
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        db.close()

@router.patch("/api/template/cell")
@router.patch("/template/cell")
async def update_template_cell(req: CellOverrideRequest):
    from core.database.db_manager import SessionLocal, Blueprint
    customer_code = req.customer_code
    locale = req.locale
    
    db = SessionLocal()
    try:
        row = db.query(Blueprint).filter(
            Blueprint.customer_code == customer_code,
            Blueprint.locale == locale
        ).first()
        
        if not row:
            return JSONResponse(status_code=404, content={"error": "Not found"})
            
        data = json.loads(row.template_json)
        sheet = data.get("template_layout", {}).get(req.sheet_name)
        if not sheet:
            return JSONResponse(status_code=404, content={"error": "Sheet not found"})
        
        def get_max_row(content, merges, styles=None):
            max_r = 0
            for addr in content.keys():
                _, r = coordinate_from_string(addr); max_r = max(max_r, r)
            m_list = merges if isinstance(merges, list) else merges.keys() if isinstance(merges, dict) else []
            for m in m_list:
                _, _, _, mr = range_boundaries(m); max_r = max(max_r, mr)
            if styles:
                for key, value in styles.items():
                    if isinstance(value, list):
                        for coord in value:
                            _, r = coordinate_from_string(coord); max_r = max(max_r, r)
                    elif isinstance(value, (dict, str)):
                        try:
                            _, r = coordinate_from_string(key); max_r = max(max_r, r)
                        except Exception:
                            pass
            return max_r

        h_content = sheet.get("template_header_content") or sheet.get("header_content", {})
        h_styles = sheet.get("template_header_styles") or sheet.get("header_styles", {})
        h_max = get_max_row(h_content, sheet.get("template_header_merges") or sheet.get("header_merges", []), styles=h_styles)
        col_letter, row_val = coordinate_from_string(req.cell_address)
        col_idx = column_index_from_string(col_letter)
        is_f = row_val > h_max

        if is_f:
            rel = row_val - h_max - 1
            f_rows = sheet.get("template_footer_rows") or sheet.get("footer_rows", [])
            row_item = next((r for r in f_rows if r.get('relative_index') == rel), None)
            if not row_item:
                row_item = {"relative_index": rel, "cells": [], "merges": []}
                f_rows.append(row_item)
                sheet["template_footer_rows"] = sorted(f_rows, key=lambda x: x.get('relative_index', 0))
            cells = row_item.get("cells", [])
            cell = next((c for c in cells if c.get('col_index') == col_idx), None)
            if not cell:
                cell = {"col_index": col_idx, "value": ""}
                cells.append(cell)
                row_item["cells"] = sorted(cells, key=lambda x: x.get('col_index', 1))
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

        row.template_json = json.dumps(data, ensure_ascii=False)
        db.commit()
        
        # Clear cache file so it regenerates from updated DB
        _clear_cache(customer_code, locale)
        
        return {"status": "success"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        db.close()

@router.patch("/template/notes")
async def update_template_notes(req: TemplateNotesRequest):
    from core.database.db_manager import SessionLocal, Blueprint
    customer_code = req.customer_code
    locale = req.locale
    
    db = SessionLocal()
    try:
        row = db.query(Blueprint).filter(
            Blueprint.customer_code == customer_code,
            Blueprint.locale == locale
        ).first()
        
        if not row:
            return JSONResponse(status_code=404, content={"error": "Not found"})
            
        data = json.loads(row.template_json)
        data["notes"] = req.notes
        row.template_json = json.dumps(data, ensure_ascii=False)
        db.commit()
        
        # Clear cache
        _clear_cache(customer_code, locale)
        
        return {"status": "success"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        db.close()

@router.delete("/template/{customer_code}")
async def delete_template(customer_code: str, locale: str = "KH"):
    from core.database.db_manager import SessionLocal, Blueprint
    
    db = SessionLocal()
    try:
        row = db.query(Blueprint).filter(
            Blueprint.customer_code == customer_code,
            Blueprint.locale == locale
        ).first()
        
        if not row:
            return JSONResponse(status_code=404, content={"error": "Not found"})
            
        db.delete(row)
        db.commit()
        
        # Clear cache
        _clear_cache(customer_code, locale)
        
        return {"status": "success"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        db.close()
