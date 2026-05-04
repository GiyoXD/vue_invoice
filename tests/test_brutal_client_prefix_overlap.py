import pytest
from pathlib import Path
from core.invoice_generator.resolvers import InvoiceAssetResolver

@pytest.fixture
def test_env(tmp_path: Path):
    config_dir = tmp_path / "config"
    template_dir = tmp_path / "templates"
    config_dir.mkdir()
    template_dir.mkdir()
    return config_dir, template_dir

def create_mock_client(config_dir: Path, folder_name: str, config_name: str = None):
    """Helper to create a dummy client folder with required files."""
    if config_name is None:
        config_name = folder_name
    folder = config_dir / folder_name
    folder.mkdir(exist_ok=True)
    (folder / f"{config_name}_config.json").touch()
    (folder / f"{config_name}.xlsx").touch()
    return folder

def test_brutal_prefix_overlaps(test_env):
    """
    Brutal testing for edge cases in prefix matching.
    We test combinations of:
    - Hyphens
    - Substrings
    - Numbers
    - Strict non-matches
    """
    config_dir, template_dir = test_env
    resolver = InvoiceAssetResolver(config_dir, template_dir)

    # 1. Setup confusingly similar clients
    create_mock_client(config_dir, "JLFTLT-VC")
    create_mock_client(config_dir, "JLFTLT")
    create_mock_client(config_dir, "JLFT")     # Substring of JLFTLT
    create_mock_client(config_dir, "JLF")      # Substring of JLFT
    create_mock_client(config_dir, "JLFTLT_INV") # Underscore suffix as different client

    # 2. Test Exact Matches (they should perfectly resolve to their own folders)
    assert "JLFTLT-VC" in str(resolver.resolve_assets_for_input_file("JLFTLT-VC25001.json").config_path)
    assert "JLFTLT_config.json" in str(resolver.resolve_assets_for_input_file("JLFTLT25001.json").config_path)
    assert "JLFT_config.json" in str(resolver.resolve_assets_for_input_file("JLFT25001.json").config_path)
    assert "JLF_config.json" in str(resolver.resolve_assets_for_input_file("JLF25001.json").config_path)

    # 3. Test Deletions (Removing exact matches should NOT cause fallback to overlapping substrings or other clients)
    
    # Remove JLFTLT folder to see if "JLFTLT25001" falls back to something else
    import shutil
    shutil.rmtree(config_dir / "JLFTLT")
    
    # JLFTLT25001 should now fail (return None). 
    # It must NOT resolve to "JLFTLT-VC", or "JLFTLT_INV".
    assets = resolver.resolve_assets_for_input_file("JLFTLT25001.json")
    assert assets is None, f"JLFTLT mistakenly matched another folder! Resolved to {assets}"

    # Remove JLFT folder to see if "JLFT25001" falls back
    shutil.rmtree(config_dir / "JLFT")
    assets = resolver.resolve_assets_for_input_file("JLFT25001.json")
    assert assets is None, f"JLFT mistakenly matched JLFTLT or JLF! Resolved to {assets}"

    # Remove JLF folder
    shutil.rmtree(config_dir / "JLF")
    assets = resolver.resolve_assets_for_input_file("JLF25001.json")
    assert assets is None, "JLF matched something else!"

def test_legacy_fallback_folder_logic(test_env):
    """
    Test that valid fallback folders (like JF_config) still work,
    but invalid overlaps (like JLFTLT_INV for JLFTLT) do not break logic.
    """
    config_dir, template_dir = test_env
    resolver = InvoiceAssetResolver(config_dir, template_dir)

    # Create a legacy config folder instead of a direct named folder
    create_mock_client(config_dir, "LEGACY_config", "LEGACY")
    
    # LEGACY25001 should match LEGACY_config folder
    assets = resolver.resolve_assets_for_input_file("LEGACY25001.json")
    assert assets is not None
    assert "LEGACY_config" in str(assets.config_path)
