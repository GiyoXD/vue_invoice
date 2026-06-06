import json
import pytest
from core.database.db_manager import Blueprint, BlueprintTemplate
from core.invoice_generator.resolvers import InvoiceAssetResolver

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
        config_json=mock_config,
        template_json=mock_template_layout
    )
    blueprint.template_binary = BlueprintTemplate(
        filename=f"{customer_code}_{locale}.xlsx",
        xlsx_blob=mock_xlsx_bytes
    )
    db.add(blueprint)
    db.commit()


def test_brutal_prefix_overlaps(db):
    """
    Brutal testing for edge cases in prefix matching: hyphens, substrings, numbers, etc.
    """
    resolver = InvoiceAssetResolver()
    test_codes = ["JLFTLT-VC", "JLFTLT", "JLFT", "JLF", "JLFTLT_INV"]
    
    for code in test_codes:
        create_db_client(db, code)

    # Exact matches should perfectly resolve to their own records
    assert resolver.resolve_assets_for_input_file("JLFTLT-VC25001.json").config_data["_meta"]["customer"] == "JLFTLT-VC_KH"
    assert resolver.resolve_assets_for_input_file("JLFTLT25001.json").config_data["_meta"]["customer"] == "JLFTLT_KH"
    assert resolver.resolve_assets_for_input_file("JLFT25001.json").config_data["_meta"]["customer"] == "JLFT_KH"
    assert resolver.resolve_assets_for_input_file("JLF25001.json").config_data["_meta"]["customer"] == "JLF_KH"

    # Removing exact matches should NOT cause fallback to overlapping substrings or other clients
    db.query(Blueprint).filter(Blueprint.customer_code == "JLFTLT").delete()
    db.commit()
    assert resolver.resolve_assets_for_input_file("JLFTLT25001.json") is None

    db.query(Blueprint).filter(Blueprint.customer_code == "JLFT").delete()
    db.commit()
    assert resolver.resolve_assets_for_input_file("JLFT25001.json") is None

    db.query(Blueprint).filter(Blueprint.customer_code == "JLF").delete()
    db.commit()
    assert resolver.resolve_assets_for_input_file("JLF25001.json") is None


def test_validate_similar_client_prefix_not_overlapse(db):
    """
    Test that sourcing correctly differentiates between similar client prefixes like JLFTLT and JLFTLT-VC.
    """
    resolver = InvoiceAssetResolver()
    
    # Create JLFTLT-VC in DB
    create_db_client(db, "JLFTLT-VC")
    
    # "JLFTLT25001.json" prefix is "JLFTLT", which should not match "JLFTLT-VC"
    assert resolver.resolve_assets_for_input_file("JLFTLT25001.json") is None
    
    # Create JLFT in DB
    create_db_client(db, "JLFTLT")
    
    # Both should now resolve to their correct folders
    assert resolver.resolve_assets_for_input_file("JLFTLT-VC25001.json").config_data["_meta"]["customer"] == "JLFTLT-VC_KH"
    assert resolver.resolve_assets_for_input_file("JLFTLT25001.json").config_data["_meta"]["customer"] == "JLFTLT_KH"


def test_differentiate_trailing_hyphen(db):
    """
    Test that sourcing correctly differentiates between prefixes with trailing hyphens (KB-) and those without (KB).
    """
    resolver = InvoiceAssetResolver()
    
    create_db_client(db, "KB-")
    create_db_client(db, "KB")
    
    assert resolver.resolve_assets_for_input_file("KB-25001.json").config_data["_meta"]["customer"] == "KB-_KH"
    assert resolver.resolve_assets_for_input_file("KB25001.json").config_data["_meta"]["customer"] == "KB_KH"
