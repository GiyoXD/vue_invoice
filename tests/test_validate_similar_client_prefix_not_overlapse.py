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

def test_validate_similar_client_prefix_not_overlapse(setup_db):
    """
    Test that the sourcing algorithm correctly differentiates between
    similar client prefixes like JLFTLT and JLFTLT-VC, preventing overlaps.
    If JLFTLT matches against JLFTLT-VC (or vice versa), it should trigger an error
    or fail to resolve instead of returning the wrong client's assets.
    """
    db = SessionLocal()
    
    test_codes = ["JLFTLT-VC", "JLFTLT"]
    db.query(Blueprint).filter(Blueprint.customer_code.in_(test_codes)).delete()
    db.commit()
    
    resolver = InvoiceAssetResolver()
    
    try:
        # Create JLFTLT-VC in DB
        create_db_client(db, "JLFTLT-VC")
        
        # If the user submits "JLFTLT25001.json", the prefix is "JLFTLT".
        # It should NOT match "JLFTLT-VC" folder.
        assets = resolver.resolve_assets_for_input_file("JLFTLT25001.json")
        assert assets is None, "Should return None since JLFTLT does not exist."
        
        # Conversely, test that JLFTLT-VC perfectly matches its own folder and not something else.
        # Create JLFTLT folder now
        create_db_client(db, "JLFTLT")
        
        # JLFTLT-VC should resolve to JLFTLT-VC
        vc_assets = resolver.resolve_assets_for_input_file("JLFTLT-VC25001.json")
        assert vc_assets is not None
        assert vc_assets.config_data["_meta"]["customer"] == "JLFTLT-VC_KH"
        
        # JLFTLT should now resolve to JLFTLT
        jlftlt_assets = resolver.resolve_assets_for_input_file("JLFTLT25001.json")
        assert jlftlt_assets is not None
        assert jlftlt_assets.config_data["_meta"]["customer"] == "JLFTLT_KH"
        
    finally:
        db.query(Blueprint).filter(Blueprint.customer_code.in_(test_codes)).delete()
        db.commit()
        db.close()


def test_differentiate_trailing_hyphen(setup_db):
    """
    Test that the sourcing algorithm correctly differentiates between
    prefixes with trailing hyphens (like KB-) and those without (like KB).
    They must not overlap.
    """
    db = SessionLocal()
    
    test_codes = ["KB-", "KB"]
    db.query(Blueprint).filter(Blueprint.customer_code.in_(test_codes)).delete()
    db.commit()
    
    resolver = InvoiceAssetResolver()
    
    try:
        # Create KB-
        create_db_client(db, "KB-")
        
        # Create KB
        create_db_client(db, "KB")
        
        # KB-25001.json has prefix "KB-", so it should resolve to KB-
        dash_assets = resolver.resolve_assets_for_input_file("KB-25001.json")
        assert dash_assets is not None
        assert dash_assets.config_data["_meta"]["customer"] == "KB-_KH"
        
        # KB25001.json has prefix "KB", so it should resolve to KB
        kb_assets = resolver.resolve_assets_for_input_file("KB25001.json")
        assert kb_assets is not None
        assert kb_assets.config_data["_meta"]["customer"] == "KB_KH"
        
    finally:
        db.query(Blueprint).filter(Blueprint.customer_code.in_(test_codes)).delete()
        db.commit()
        db.close()
