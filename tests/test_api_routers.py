import json
import pytest
from core.database.db_manager import Blueprint, BlueprintTemplate

def test_api_template_flow(client, db):
    # 1. Clean up any existing test records first
    db.query(Blueprint).filter(Blueprint.customer_code == "APITEST").delete()
    db.commit()

    # 2. View non-existent template
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 404

    # 3. Insert mock blueprint directly into DB so we can test view, patch, and delete
    mock_config = {
        "_meta": {
            "config_version": "2.2_strict_mode",
            "customer": "APITEST_KH",
            "description": "Mock API Test Config"
        },
        "table_info": {
            "fallback_description": {"standard": "Mock Standard Desc"}
        }
    }
    mock_template = {
        "fingerprint": {
            "source_file": "APITEST_raw.xlsx"
        },
        "template_layout": {
            "Invoice": {
                "header_rows": [
                    {
                        "relative_index": 0,
                        "height": 20.0,
                        "cells": [
                            {
                                "col_index": 1,
                                "value": "Original Value"
                            }
                        ]
                    }
                ],
                "footer_rows": [],
                "col_widths": {}
            }
        }
    }
    blueprint = Blueprint(
        customer_code="APITEST",
        locale="KH",
        description="API Test Blueprint",
        config_json=mock_config,
        template_json=mock_template
    )
    blueprint.template_binary = BlueprintTemplate(
        filename="APITEST_KH.xlsx",
        xlsx_blob=b"MockXlsxBytes"
    )
    db.add(blueprint)
    db.commit()

    # 4. List templates and verify the response contains customer_code and locale
    response = client.get("/api/templates")
    assert response.status_code == 200
    templates_list = response.json()
    api_test_tmpl = next((t for t in templates_list if t["name"] == "APITEST_KH"), None)
    assert api_test_tmpl is not None
    assert api_test_tmpl["customer_code"] == "APITEST"
    assert api_test_tmpl["locale"] == "KH"

    # 5. View template details
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 200
    data = response.json()
    sheet = data["template_layout"]["Invoice"]
    row0 = next(r for r in sheet["header_rows"] if r["relative_index"] == 0)
    cell_a1 = next(c for c in row0["cells"] if c["col_index"] == 1)
    assert cell_a1["value"] == "Original Value"

    # 6. Patch template cell overrides
    response = client.patch("/api/template/cell", json={
        "customer_code": "APITEST",
        "locale": "KH",
        "sheet_name": "Invoice",
        "cell_address": "A1",
        "overrides": {"standard": "New Overridden Value"}
    })
    assert response.status_code == 200

    # 7. Verify cell patch was saved
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 200
    data = response.json()
    sheet = data["template_layout"]["Invoice"]
    row0 = next(r for r in sheet["header_rows"] if r["relative_index"] == 0)
    cell_a1 = next(c for c in row0["cells"] if c["col_index"] == 1)
    cell_val = cell_a1["value"]
    assert isinstance(cell_val, dict)
    assert cell_val["standard"] == "New Overridden Value"

    # 8. Patch template notes
    response = client.patch("/api/template/notes", json={
        "customer_code": "APITEST",
        "locale": "KH",
        "notes": "Test Client Notes"
    })
    assert response.status_code == 200

    # 9. Verify notes were saved
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 200
    data = response.json()
    assert data["notes"] == "Test Client Notes"

    # 10. Delete template
    response = client.delete("/api/template/APITEST?locale=KH")
    assert response.status_code == 200

    # 11. Verify template deleted
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 404


def test_api_source_folder(client, db, tmp_path):
    # 1. Test GET default source folder
    response = client.get("/api/source-folder")
    assert response.status_code == 200
    data = response.json()
    assert "folder_path" in data

    # 2. Test POST empty / whitespace folder
    res_empty = client.post("/api/source-folder", json={"folder_path": ""})
    assert res_empty.status_code == 400
    assert "Folder path cannot be empty" in res_empty.json()["error"]

    res_ws = client.post("/api/source-folder", json={"folder_path": "   "})
    assert res_ws.status_code == 400
    assert "Folder path cannot be empty" in res_ws.json()["error"]

    # 3. Test POST invalid non-existent folder
    invalid_path = str(tmp_path / "non_existent_dir_12345")
    response = client.post("/api/source-folder", json={"folder_path": invalid_path})
    assert response.status_code == 400

    # 4. Test POST valid folder
    valid_folder = tmp_path / "custom_sources"
    valid_folder.mkdir(parents=True, exist_ok=True)
    response = client.post("/api/source-folder", json={"folder_path": str(valid_folder)})
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "success"
    assert data["folder_path"] == str(valid_folder.resolve())

    # 5. Test GET updated folder
    response = client.get("/api/source-folder")
    assert response.status_code == 200
    assert response.json()["folder_path"] == str(valid_folder.resolve())


def test_api_source_files(client, db, tmp_path):
    # 1. Setup custom source folder
    source_dir = tmp_path / "source_files_dir"
    source_dir.mkdir(parents=True, exist_ok=True)
    client.post("/api/source-folder", json={"folder_path": str(source_dir)})

    # 2. Create sample files
    f1 = source_dir / "sample_a.xlsx"
    f1.write_bytes(b"dummy excel A")
    f2 = source_dir / "sample_b.xls"
    f2.write_bytes(b"dummy excel B")
    f_temp = source_dir / "~$sample_temp.xlsx"
    f_temp.write_bytes(b"temp excel")
    f_txt = source_dir / "ignore.txt"
    f_txt.write_text("ignore me")

    # 3. Call GET /api/source-files
    response = client.get("/api/source-files")
    assert response.status_code == 200
    data = response.json()
    assert "files" in data
    assert "folder_path" in data
    filenames = [f["filename"] for f in data["files"]]
    assert "sample_a.xlsx" in filenames
    assert "sample_b.xls" in filenames
    assert "~$sample_temp.xlsx" not in filenames
    assert "ignore.txt" not in filenames
    assert len(data["files"]) == 2


@pytest.fixture
def sample_xlsx_path(tmp_path):
    from pathlib import Path
    import shutil
    import openpyxl

    sample_file = Path("tests/experiment_sample/shipping_list/JF25057.xlsx")
    dest_path = tmp_path / "JF25057.xlsx"
    if sample_file.exists():
        shutil.copy2(sample_file, dest_path)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["PO", "Item", "Description", "Qty", "SF", "Price", "Amount", "Gross", "Net", "CBM"])
        ws.append(["PO001", "ITM01", "Standard item", 10, 100.0, 5.0, 500.0, 50.0, 45.0, "1x1x1"])
        wb.save(dest_path)
    return dest_path


def test_api_security_path_traversal_and_extensions(client, db, tmp_path):
    source_dir = tmp_path / "sec_source_dir"
    source_dir.mkdir(parents=True, exist_ok=True)
    client.post("/api/source-folder", json={"folder_path": str(source_dir)})

    # 1. Reject invalid extensions for /api/process-existing
    res = client.post("/api/process-existing", json={"filename": "test.exe"})
    assert res.status_code == 400
    assert "Only .xlsx and .xls files are supported" in res.json()["error"]

    res = client.post("/api/process-existing", json={"filename": "../../malicious.exe"})
    assert res.status_code == 400
    assert "Only .xlsx and .xls files are supported" in res.json()["error"]

    # 2. Path traversal sanitized for /api/process-existing (missing file returns 404)
    res = client.post("/api/process-existing", json={"filename": "../../malicious.xlsx"})
    assert res.status_code == 404
    assert res.json()["error"] == "File not found: malicious.xlsx"

    # 3. Reject invalid extensions for /api/open-file
    res = client.post("/api/open-file", json={"filename": "test.exe"})
    assert res.status_code == 400
    assert "Only .xlsx and .xls files are supported" in res.json()["error"]

    res = client.post("/api/open-file", json={"filename": "../../malicious.exe"})
    assert res.status_code == 400
    assert "Only .xlsx and .xls files are supported" in res.json()["error"]

    # 4. Path traversal sanitized for /api/open-file
    res = client.post("/api/open-file", json={"filename": "../../malicious.xlsx"})
    assert res.status_code == 404
    assert res.json()["error"] == "File not found: malicious.xlsx"

    # 5. Reject invalid extensions for /api/upload
    res = client.post(
        "/api/upload",
        files={"file": ("malicious.exe", b"MZbinarycontent", "application/octet-stream")}
    )
    assert res.status_code == 400
    assert "Only .xlsx and .xls files are supported" in res.json()["error"]

    res = client.post(
        "/api/upload",
        files={"file": ("script.sh", b"#!/bin/bash\necho hello", "text/plain")}
    )
    assert res.status_code == 400
    assert "Only .xlsx and .xls files are supported" in res.json()["error"]


def test_api_process_existing_and_open(client, db, tmp_path, monkeypatch, sample_xlsx_path):
    import shutil

    source_dir = tmp_path / "source_proc_dir"
    source_dir.mkdir(parents=True, exist_ok=True)
    client.post("/api/source-folder", json={"folder_path": str(source_dir)})

    # Copy fixture sample file to source directory
    target_name = sample_xlsx_path.name
    shutil.copy2(sample_xlsx_path, source_dir / target_name)

    # 1. Test process non-existent file
    response = client.post("/api/process-existing", json={"filename": "missing.xlsx"})
    assert response.status_code == 404

    # 2. Test process existing file
    response = client.post("/api/process-existing", json={"filename": target_name})
    assert response.status_code in [200, 422]
    if response.status_code == 200:
        data = response.json()
        assert data["status"] == "success"
        assert data["file_name"] == target_name
        assert "identifier" in data
        assert "json_path" in data

    # 3. Test open-file
    opened_paths = []
    monkeypatch.setattr("os.startfile", lambda p: opened_paths.append(p), raising=False)
    monkeypatch.setattr("subprocess.Popen", lambda args: opened_paths.append(args), raising=False)

    # 4. Open non-existent file
    response = client.post("/api/open-file", json={"filename": "missing.xlsx"})
    assert response.status_code == 404

    # 5. Open existing file
    response = client.post("/api/open-file", json={"filename": target_name})
    assert response.status_code == 200
    assert response.json()["status"] == "success"
    assert len(opened_paths) == 1

