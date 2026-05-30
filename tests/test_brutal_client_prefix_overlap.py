import json
import pytest
from core.database.db_manager import init_db, SessionLocal, Blueprint, BlueprintTemplate
from core.invoice_generator.resolvers import InvoiceAssetResolver

@pytest.fixture(scope="module")
def setup_db():
    init_db()
    yield

def create_db_client(db, customer_code: str, locale: str = "KH"):
    """Helper to create a dummy blueprint in the database."""
    mock_config = {
        "_meta": {
            "config_version": "2.2_strict_mode",
            "customer": f"{customer_code}_{locale}",
            "description": f"Mock config for {customer_code}"
        },
        "processing": {
            "sheets": ["Invoice"]
        }
    }
    mock_template_layout = {
        "fingerprint": {
            "source_file": f"{customer_code}_raw.xlsx"
        },
        "template_layout": {
            "Invoice": {
                "header_content": {}
            }
        }
    }
    mock_xlsx_bytes = b"PK\x03\x04MockExcelTemplateBinaryDataBytes"
    
    blueprint = Blueprint(
        customer_code=customer_code,
        locale=locale,
        description=f"Test blueprint for {customer_code}",
        config_json=json.dumps(mock_config),
        template_json=json.dumps(mock_template_layout)
    )
    blueprint.template_binary = BlueprintTemplate(
        filename=f"{customer_code}_{locale}.xlsx",
        xlsx_blob=mock_xlsx_bytes
    )
    db.add(blueprint)
    db.commit()

def test_brutal_prefix_overlaps(setup_db):
    """
    Brutal testing for edge cases in prefix matching.
    We test combinations of:
    - Hyphens
    - Substrings
    - Numbers
    - Strict non-matches
    """
    db = SessionLocal()
    
    # Clean up any existing test records
    test_codes = ["JLFTLT-VC", "JLFTLT", "JLFT", "JLF", "JLFTLT_INV"]
    db.query(Blueprint).filter(Blueprint.customer_code.in_(test_codes)).delete()
    db.commit()

    resolver = InvoiceAssetResolver()

    try:
        # 1. Setup confusingly similar clients in DB
        for code in test_codes:
            create_db_client(db, code)

        # 2. Test Exact Matches (they should perfectly resolve to their own records)
        assert resolver.resolve_assets_for_input_file("JLFTLT-VC25001.json").config_data["_meta"]["customer"] == "JLFTLT-VC_KH"
        assert resolver.resolve_assets_for_input_file("JLFTLT25001.json").config_data["_meta"]["customer"] == "JLFTLT_KH"
        assert resolver.resolve_assets_for_input_file("JLFT25001.json").config_data["_meta"]["customer"] == "JLFT_KH"
        assert resolver.resolve_assets_for_input_file("JLF25001.json").config_data["_meta"]["customer"] == "JLF_KH"

        # 3. Test Deletions (Removing exact matches should NOT cause fallback to overlapping substrings or other clients)
        
        # Remove JLFTLT from DB to see if "JLFTLT25001" falls back to something else
        db.query(Blueprint).filter(Blueprint.customer_code == "JLFTLT").delete()
        db.commit()
        
        # JLFTLT25001 should now fail (return None). 
        # It must NOT resolve to "JLFTLT-VC", or "JLFTLT_INV".
        assets = resolver.resolve_assets_for_input_file("JLFTLT25001.json")
        assert assets is None, f"JLFTLT mistakenly matched another record! Resolved to {assets}"

        # Remove JLFT
        db.query(Blueprint).filter(Blueprint.customer_code == "JLFT").delete()
        db.commit()
        assets = resolver.resolve_assets_for_input_file("JLFT25001.json")
        assert assets is None, f"JLFT mistakenly matched JLFTLT or JLF!"

        # Remove JLF
        db.query(Blueprint).filter(Blueprint.customer_code == "JLF").delete()
        db.commit()
        assets = resolver.resolve_assets_for_input_file("JLF25001.json")
        assert assets is None, "JLF matched something else!"

    finally:
        db.query(Blueprint).filter(Blueprint.customer_code.in_(test_codes)).delete()
        db.commit()
        db.close()
