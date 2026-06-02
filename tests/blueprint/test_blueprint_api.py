import io
import json
import pytest
import openpyxl
from io import BytesIO
from core.database.db_manager import Blueprint

def test_blueprint_scan_and_generate_flow(client, db):
    # 1. Create a minimal valid Excel template in memory that matches scanner expectations
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice"
    
    # Headers at Row 3
    ws.cell(row=3, column=1, value="Mark & No")
    ws.cell(row=3, column=2, value="P.O. No.")
    ws.cell(row=3, column=3, value="Random Unmapped Column")  # Unknown header to trigger 'needs_mapping'
    ws.cell(row=3, column=4, value="Quantity")
    ws.cell(row=3, column=5, value="Unit Price")
    ws.cell(row=3, column=6, value="Amount")
    
    # Description Fallback in Column 1, Row 4 (part of static column values)
    ws.cell(row=4, column=1, value="DES: COW LEATHER")
    ws.cell(row=4, column=3, value="Product A")
    ws.cell(row=4, column=4, value=10)
    ws.cell(row=4, column=5, value=5.0)
    ws.cell(row=4, column=6, value=50.0)
    
    # HS Code in Row 8
    ws.cell(row=8, column=5, value="HS CODE: 4107.12.00")
    
    # Footer Row at Row 10
    ws.cell(row=10, column=2, value="TOTAL:")
    ws.cell(row=10, column=6, value=50.0)
    
    buf = BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    # 2. Upload the file to Step 1: /api/blueprint/scan
    response = client.post(
        "/api/blueprint/scan",
        files={"file": ("test_api_template.xlsx", xlsx_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    )
    assert response.status_code == 200
    scan_data = response.json()
    
    assert scan_data["status"] == "needs_mapping"
    assert scan_data["file_token"] == "scan_test_api_template.xlsx"
    assert "Random Unmapped Column" in scan_data["unknown_headers"]

    # 3. Post user mappings to Step 2: /api/blueprint/generate
    generate_payload = {
        "file_token": scan_data["file_token"],
        "customer_code": "APITESTGEN",
        "locale": "KH",
        "mappings": {
            "P.O. No.": "col_po",
            "Random Unmapped Column": "col_desc",
            "Quantity": "col_qty_sf",
            "Unit Price": "col_unit_price",
            "Amount": "col_amount"
        },
        "footer_mappings": [],
        "pricing_mode": "standard"
    }
    
    response = client.post("/api/blueprint/generate", json=generate_payload)
    assert response.status_code == 200
    gen_data = response.json()
    assert gen_data["status"] == "success"
    assert "generated and saved to database" in gen_data["message"]

    # 4. Verify that the generated template is saved in the database and readable
    response = client.get("/api/template/view?customer_code=APITESTGEN&locale=KH")
    assert response.status_code == 200
    template_data = response.json()
    assert "template_layout" in template_data
    assert "Invoice" in template_data["template_layout"]
    


def test_deliberate_failure_for_demo():
    """A deliberate failing test to verify that pytest is running and reporting failures correctly."""
    assert False, "This is a deliberate failure to show that pytest is executing and reporting errors."

