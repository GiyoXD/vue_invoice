import logging
import os
import re
import traceback
from pathlib import Path
from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import Optional
from datetime import datetime
from core.system_config import sys_config

_DEBUG = os.getenv("GIYO_DEBUG", "").strip() == "1"

try:
    import gspread
    from google.oauth2.service_account import Credentials
    _GSPREAD_AVAILABLE = True
except ImportError:
    _GSPREAD_AVAILABLE = False

router = APIRouter(prefix="/api/sheets", tags=["google_sheets"])
logger = logging.getLogger(__name__)

# Single source of truth for the default Google Spreadsheet ID.
# Override via the UI field or the GOOGLE_SHEET_ID env var.
DEFAULT_SPREADSHEET_ID = "1piQv1mpEWWBw3hUITAD2Q5ZHdvnqP38n0OLUbK5Y1Hk"

class LookupRequest(BaseModel):
    invoice_no: str
    spreadsheet_id: Optional[str] = None
    worksheet_name: Optional[str] = "2026"

class SheetDataPayload(BaseModel):
    invoice_no: str
    ref_no: Optional[str] = ""
    invoice_date: str
    summary: Optional[str] = ""

class ExportRequest(BaseModel):
    payload: SheetDataPayload
    spreadsheet_id: Optional[str] = None
    worksheet_name: Optional[str] = "2026"
    force_override: bool = False



def _get_worksheet(spreadsheet_id: Optional[str], worksheet_name: str):
    if not _GSPREAD_AVAILABLE:
        raise ValueError("Required libraries missing. Please run `pip install gspread google-auth`")

    # 1. Resolve Credentials
    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not credentials_path:
        possible_paths = [
            sys_config.base_dir / "secret.json",
            sys_config.base_dir / "credentials.json",
            sys_config.base_dir / "database" / "credentials.json"
        ]
        for p in possible_paths:
            if p.exists():
                credentials_path = str(p)
                break
                
    if not credentials_path or not os.path.exists(credentials_path):
        raise ValueError("Missing Google Service Account credentials. Place 'secret.json' in project root.")

    # 2. Resolve Spreadsheet ID (UI override > env var > module constant)
    sheet_id = spreadsheet_id or os.getenv("GOOGLE_SHEET_ID") or DEFAULT_SPREADSHEET_ID

    # 3. Authenticate and Connect
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    client = gspread.authorize(credentials)

    # 4. Open Spreadsheet & Worksheet
    spreadsheet = client.open_by_key(sheet_id)
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.sheet1
        
    return worksheet

@router.post("/resolve_ref")
async def resolve_reference(request: LookupRequest):
    """
    Look up an Invoice No. 
    If found, return the existing Ref No.
    If not found, find the last Ref No in Col D, increment it, and return the new one.
    """
    try:
        worksheet = _get_worksheet(request.spreadsheet_id, request.worksheet_name)
        
        # Col C is index 3, Col D is index 4
        col_c_values = worksheet.col_values(3)
        col_d_values = worksheet.col_values(4)
        
        # Scenario 1: Invoice found — check for duplicates
        matching_indices = [i for i, v in enumerate(col_c_values) if v == request.invoice_no]
        if matching_indices:
            # Use the LAST match (most recent row), not the first
            last_match_idx = matching_indices[-1]
            ref_no = ""
            if len(col_d_values) > last_match_idx:
                ref_no = col_d_values[last_match_idx]
            
            result = {"found": True, "ref_no": ref_no or ""}
            if len(matching_indices) > 1:
                result["warning"] = f"Duplicate: invoice '{request.invoice_no}' appears in {len(matching_indices)} rows. Using the latest (row {last_match_idx + 1})."
                logger.warning(f"Duplicate invoice '{request.invoice_no}' found in rows: {[i+1 for i in matching_indices]}")
            return result
        
        # Scenario 2: Invoice not found, generate next ref no.
        # Scan ALL refs and use max numeric suffix to avoid collisions
        # from out-of-order manual inserts.
        max_num = 0
        max_num_width = 3  # default zero-pad width
        ref_prefix = ""
        
        for val in col_d_values:
            val_str = str(val).strip() if val else ""
            if not val_str or val_str == "Ref No":  # skip blanks and header
                continue
            match = re.search(r'(\d+)$', val_str)
            if match:
                num_str = match.group(1)
                num_val = int(num_str)
                if num_val >= max_num:
                    max_num = num_val
                    max_num_width = len(num_str)
                    ref_prefix = val_str[:match.start()]
        
        if ref_prefix:
            next_num = max_num + 1
            new_num_str = f"{next_num:0{max_num_width}d}"
            next_ref = ref_prefix + new_num_str
        else:
            next_ref = "REF-001"  # Default fallback (empty sheet)
            
        return {"found": False, "ref_no": next_ref}

    except ValueError as ve:
        return JSONResponse(status_code=400, content={"error": str(ve)})
    except Exception as e:
        logger.error(f"Failed to resolve ref sheet: {str(e)}\n{traceback.format_exc()}")
        content = {"error": str(e)}
        if _DEBUG:
            content["traceback"] = traceback.format_exc()
        return JSONResponse(status_code=500, content=content)

@router.post("/export")
async def export_to_google_sheets(request: ExportRequest):
    """
    Export or update invoice details based on matching logic.
    """
    try:
        data = request.payload
        worksheet = _get_worksheet(request.spreadsheet_id, request.worksheet_name)
        col_c_values = worksheet.col_values(3)
        col_d_values = worksheet.col_values(4)
        
        # 1. Duplicate Ref No Check
        ref_no_clean = data.ref_no.strip() if data.ref_no else ""
        if ref_no_clean and ref_no_clean in col_d_values:
            ref_idx = col_d_values.index(ref_no_clean)
            existing_invoice_for_ref = col_c_values[ref_idx] if ref_idx < len(col_c_values) else ""
            
            # If the Ref No exists but belongs to a completely different Invoice, block it!
            if existing_invoice_for_ref and existing_invoice_for_ref != data.invoice_no:
                return JSONResponse(status_code=400, content={
                    "error": f"Duplicate Ref No: '{ref_no_clean}' is already assigned to Invoice '{existing_invoice_for_ref}'. Please use a different Reference Number."
                })
        
        try:
            dt = datetime.strptime(data.invoice_date, "%Y-%m-%d")
            formatted_date = dt.strftime("%d/%m/%Y")
        except Exception:
            formatted_date = data.invoice_date

        combined_val = data.summary or ""

        if data.invoice_no in col_c_values:
            if not request.force_override:
                return {
                    "status": "conflict",
                    "message": "Invoice already exists in spreadsheet. Do you want to override it?",
                    "action": "conflict"
                }
            else:
                row_idx = col_c_values.index(data.invoice_no) + 1
                updates = [
                    {'range': f'C{row_idx}', 'values': [[data.invoice_no]]},
                    {'range': f'D{row_idx}', 'values': [[data.ref_no]]},
                    {'range': f'F{row_idx}', 'values': [[formatted_date]]},
                    {'range': f'M{row_idx}', 'values': [[combined_val]]}
                ]
                worksheet.batch_update(updates, value_input_option='USER_ENTERED')
                logger.info(f"Successfully overridden invoice {data.invoice_no} safely using batch update.")
                return {
                    "status": "success",
                    "message": "Data successfully overridden without affecting other columns.",
                    "action": "updated"
                }
        else:
            # Determine next row using Column C (Invoice No) as the single source of truth for row existence.
            next_row = len(col_c_values) + 1
            updates = [
                {'range': f'C{next_row}', 'values': [[data.invoice_no]]},
                {'range': f'D{next_row}', 'values': [[data.ref_no]]},
                {'range': f'F{next_row}', 'values': [[formatted_date]]},
                {'range': f'M{next_row}', 'values': [[combined_val]]}
            ]
            worksheet.batch_update(updates, value_input_option='USER_ENTERED')
            logger.info(f"Successfully exported new invoice {data.invoice_no} to Google Sheets (row {next_row}).")
            return {
                "status": "success",
                "message": "Data successfully exported as a new row.",
                "action": "appended"
            }

    except ValueError as ve:
        return JSONResponse(status_code=400, content={"error": str(ve)})
    except Exception as e:
        logger.error(f"Failed to export to Google Sheets: {str(e)}\n{traceback.format_exc()}")
        content = {"error": str(e)}
        if _DEBUG:
            content["traceback"] = traceback.format_exc()
        return JSONResponse(status_code=500, content=content)
