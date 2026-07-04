from types import SimpleNamespace
from core.invoice_generator.utils.workbook_utils import (
    _build_split_filename,
    get_mode_suffix,
    count_layout_columns,
)

# --- Tests for _build_split_filename ---

def test_build_split_filename_replaces_invoice_word():
    # 1. Arrange
    base_stem = "TH26001_Invoice_KH"
    sheet_name = "Packing list"
    suffix = ".xlsx"
    expected = "TH26001_Packing list_KH.xlsx"

    # 2. Act
    result = _build_split_filename(base_stem, sheet_name, suffix)

    # 3. Assert
    assert result == expected


def test_build_split_filename_appends_when_no_invoice_word():
    # 1. Arrange
    base_stem_no_invoice = "Monthly_Summary"
    sheet_name = "Contract"
    suffix = ".xlsx"
    expected = "Monthly_Summary_Contract.xlsx"

    # 2. Act
    result = _build_split_filename(base_stem_no_invoice, sheet_name, suffix)

    # 3. Assert
    assert result == expected


# --- Tests for get_mode_suffix (Mocking variables) ---

def test_get_mode_suffix_daf_mode():
    # 1. Arrange
    ctx = SimpleNamespace(daf_mode=True, custom_mode=False)

    # 2. Act
    result = get_mode_suffix(ctx)

    # 3. Assert
    assert result == " DAF"


def test_get_mode_suffix_custom_mode():
    # 1. Arrange
    ctx = SimpleNamespace(daf_mode=False, custom_mode=True)

    # 2. Act
    result = get_mode_suffix(ctx)

    # 3. Assert
    assert result == " Custom"


def test_get_mode_suffix_standard_mode():
    # 1. Arrange
    ctx = SimpleNamespace(daf_mode=False, custom_mode=False)

    # 2. Act
    result = get_mode_suffix(ctx)

    # 3. Assert
    assert result == ""


# --- Tests for count_layout_columns (Mocking class methods) ---

def test_count_layout_columns_with_children():
    # 1. Arrange
    mock_layout = {
        'structure': {
            'columns': [
                {'id': 'col_po'},
                {'id': 'col_qty_header', 'children': [{'id': 'col_qty_pcs'}, {'id': 'col_qty_sf'}]},
                {'id': 'col_amount'}
            ]
        }
    }

    class DummyConfigLoader:
        def get_layout_config(self, sheet_name):
            assert sheet_name == "Invoice"
            return mock_layout

    loader = DummyConfigLoader()

    # 2. Act
    result = count_layout_columns(loader, "Invoice")

    # 3. Assert
    assert result == 4  # col_po (1) + 2 children + col_amount (1) = 4


def test_count_layout_columns_no_layout():
    # 1. Arrange
    class DummyConfigLoader:
        def get_layout_config(self, sheet_name):
            return {}  # empty layout config

    loader = DummyConfigLoader()

    # 2. Act
    result = count_layout_columns(loader, "Invoice")

    # 3. Assert
    assert result is None
