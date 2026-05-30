# tests/test_api_routers.py
import json
import pytest
from fastapi.testclient import TestClient
from api.main import app
from core.database.db_manager import SessionLocal, Blueprint, BlueprintTemplate

client = TestClient(app)

@pytest.fixture(autouse=True)
def clean_db():
    db = SessionLocal()
    db.query(Blueprint).filter(Blueprint.customer_code == "APITEST").delete()
    db.commit()
    db.close()
    yield
    db = SessionLocal()
    db.query(Blueprint).filter(Blueprint.customer_code == "APITEST").delete()
    db.commit()
    db.close()

def test_api_template_flow():
    # 1. View non-existent template
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 404

    # 2. Insert mock blueprint directly into DB so we can test view, patch, and delete
    db = SessionLocal()
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
                "header_content": {"A1": "Original Value"},
                "header_styles": {}
            }
        }
    }
    blueprint = Blueprint(
        customer_code="APITEST",
        locale="KH",
        description="API Test Blueprint",
        config_json=json.dumps(mock_config),
        template_json=json.dumps(mock_template)
    )
    blueprint.template_binary = BlueprintTemplate(
        filename="APITEST_KH.xlsx",
        xlsx_blob=b"MockXlsxBytes"
    )
    db.add(blueprint)
    db.commit()
    db.close()

    # 3. List templates and verify the response contains customer_code and locale
    response = client.get("/api/templates")
    assert response.status_code == 200
    templates_list = response.json()
    api_test_tmpl = next((t for t in templates_list if t["name"] == "APITEST_KH"), None)
    assert api_test_tmpl is not None
    assert api_test_tmpl["customer_code"] == "APITEST"
    assert api_test_tmpl["locale"] == "KH"

    # 4. View template details
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 200
    data = response.json()
    assert data["template_layout"]["Invoice"]["header_content"]["A1"] == "Original Value"

    # 5. Patch template cell overrides
    response = client.patch("/api/template/cell", json={
        "customer_code": "APITEST",
        "locale": "KH",
        "sheet_name": "Invoice",
        "cell_address": "A1",
        "overrides": {"standard": "New Overridden Value"}
    })
    assert response.status_code == 200

    # 6. Verify cell patch was saved
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 200
    data = response.json()
    cell_val = data["template_layout"]["Invoice"]["header_content"]["A1"]
    assert isinstance(cell_val, dict)
    assert cell_val["standard"] == "New Overridden Value"

    # 7. Patch template notes
    response = client.patch("/api/template/notes", json={
        "customer_code": "APITEST",
        "locale": "KH",
        "notes": "Test Client Notes"
    })
    assert response.status_code == 200

    # 8. Verify notes were saved
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 200
    data = response.json()
    assert data["notes"] == "Test Client Notes"

    # 9. Delete template
    response = client.delete("/api/template/APITEST?locale=KH")
    assert response.status_code == 200

    # 10. Verify template deleted
    response = client.get("/api/template/view?customer_code=APITEST&locale=KH")
    assert response.status_code == 404
