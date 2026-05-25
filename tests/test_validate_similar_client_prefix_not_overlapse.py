import pytest
from pathlib import Path
from core.invoice_generator.resolvers import InvoiceAssetResolver

def test_validate_similar_client_prefix_not_overlapse(tmp_path: Path):
    """
    Test that the sourcing algorithm correctly differentiates between
    similar client prefixes like JLFTLT and JLFTLT-VC, preventing overlaps.
    If JLFTLT matches against JLFTLT-VC (or vice versa), it should trigger an error
    or fail to resolve instead of returning the wrong client's assets.
    """
    config_dir = tmp_path / "config"
    template_dir = tmp_path / "templates"
    
    config_dir.mkdir()
    template_dir.mkdir()
    
    # Create JLFTLT-VC folder and assets
    vc_folder = config_dir / "JLFTLT-VC"
    vc_folder.mkdir()
    (vc_folder / "JLFTLT-VC_config.json").touch()
    (vc_folder / "JLFTLT-VC.xlsx").touch()
    
    # We DO NOT create a JLFTLT folder to test the fallback logic.
    # If the user submits "JLFTLT25001.json", the prefix is "JLFTLT".
    # It should NOT match "JLFTLT-VC" folder.
    
    resolver = InvoiceAssetResolver(config_dir, template_dir)
    
    # If it incorrectly matches, it will return assets from JLFTLT-VC folder.
    # It should either return None, or raise an error.
    assets = resolver.resolve_assets_for_input_file("JLFTLT25001.json")
    
    # Ensure it does not resolve to JLFTLT-VC
    if assets is not None:
        assert "JLFTLT-VC" not in str(assets.config_path), (
            f"Error: JLFTLT incorrectly matched against JLFTLT-VC config! Resolved to: {assets.config_path}"
        )
        # Or if the requirement is explicitly to trigger an error:
        pytest.fail("Expected an error or None, but it successfully resolved overlapping assets.")
        
    assert assets is None, "Should return None since JLFTLT does not exist."
    
    # Conversely, test that JLFTLT-VC perfectly matches its own folder and not something else.
    # Create JLFTLT folder now
    jlftlt_folder = config_dir / "JLFTLT"
    jlftlt_folder.mkdir()
    (jlftlt_folder / "JLFTLT_config.json").touch()
    (jlftlt_folder / "JLFTLT.xlsx").touch()
    
    # JLFTLT-VC should resolve to JLFTLT-VC
    vc_assets = resolver.resolve_assets_for_input_file("JLFTLT-VC25001.json")
    assert vc_assets is not None
    assert "JLFTLT-VC" in str(vc_assets.config_path)
    
    # JLFTLT should now resolve to JLFTLT
    jlftlt_assets = resolver.resolve_assets_for_input_file("JLFTLT25001.json")
    assert jlftlt_assets is not None
    assert "JLFTLT_config.json" in str(jlftlt_assets.config_path)


def test_differentiate_trailing_hyphen(tmp_path: Path):
    """
    Test that the sourcing algorithm correctly differentiates between
    prefixes with trailing hyphens (like KB-) and those without (like KB).
    They must not overlap.
    """
    config_dir = tmp_path / "config"
    template_dir = tmp_path / "templates"
    
    config_dir.mkdir()
    template_dir.mkdir()
    
    # Create KB- folder and assets
    kb_dash_folder = config_dir / "KB-"
    kb_dash_folder.mkdir()
    (kb_dash_folder / "KB-_config.json").touch()
    (kb_dash_folder / "KB-.xlsx").touch()
    
    # Create KB folder and assets
    kb_folder = config_dir / "KB"
    kb_folder.mkdir()
    (kb_folder / "KB_config.json").touch()
    (kb_folder / "KB.xlsx").touch()
    
    resolver = InvoiceAssetResolver(config_dir, template_dir)
    
    # KB-25001.json has prefix "KB-", so it should resolve to the KB- folder and assets
    dash_assets = resolver.resolve_assets_for_input_file("KB-25001.json")
    assert dash_assets is not None
    assert "KB-_config.json" in str(dash_assets.config_path)
    
    # KB25001.json has prefix "KB", so it should resolve to the KB folder and assets
    kb_assets = resolver.resolve_assets_for_input_file("KB25001.json")
    assert kb_assets is not None
    assert "KB_config.json" in str(kb_assets.config_path)

