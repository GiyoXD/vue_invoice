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
                "header_content": {"A1": "Original Value"},
                "header_styles": {}
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
    assert data["template_layout"]["Invoice"]["header_content"]["A1"] == "Original Value"

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
    cell_val = data["template_layout"]["Invoice"]["header_content"]["A1"]
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
