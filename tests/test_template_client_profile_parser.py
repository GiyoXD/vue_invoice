"""
Tests for TemplateClientProfileParser.

Covers:
  - Unit tests with synthetic header_content dicts (no file I/O)
  - Integration smoke tests against every real *_template.json in
    database/blueprints/bundled/ — verifying the parser runs without
    crashing and returns sensible types.
"""

import json
import unittest
from pathlib import Path
from typing import Any, Dict, Optional

from core.invoice_generator.extractors.template_client_profile_parser import (
    TemplateClientProfileParser,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

BUNDLED_DIR = Path(__file__).resolve().parent.parent / "database" / "blueprints" / "bundled"


def _load_header_content(template_path: Path) -> Optional[Dict[str, Any]]:
    """Load header_content from the Invoice sheet of a *_template.json file."""
    with open(template_path, encoding="utf-8") as f:
        data = json.load(f)
    layout = data.get("template_layout", {})
    invoice = layout.get("Invoice", {})
    return invoice.get("template_header_content") or invoice.get("header_content")


def _all_template_files():
    return sorted(BUNDLED_DIR.rglob("*_template.json"))


# ---------------------------------------------------------------------------
# Unit tests — synthetic data
# ---------------------------------------------------------------------------

class TestGetClientFullname(unittest.TestCase):

    def test_label_and_value_are_separate_cells(self):
        hc = {"A13": "CONSIGNEE :", "B13": "Wanek Furniture Co., LTD."}
        p = TemplateClientProfileParser(hc)
        self.assertEqual(p.get_client_fullname(), "Wanek Furniture Co., LTD.")

    def test_inline_label_value_in_same_cell(self):
        hc = {"A13": "Consignee : Acme Corp"}
        p = TemplateClientProfileParser(hc)
        self.assertEqual(p.get_client_fullname(), "Acme Corp")

    def test_case_insensitive_label(self):
        hc = {"A5": "consignee", "B5": "SomeName Ltd"}
        p = TemplateClientProfileParser(hc)
        self.assertEqual(p.get_client_fullname(), "SomeName Ltd")

    def test_missing_consignee_returns_none(self):
        hc = {"A1": "INVOICE", "B1": "No 123"}
        p = TemplateClientProfileParser(hc)
        self.assertIsNone(p.get_client_fullname())

    def test_override_dict_value_uses_default(self):
        hc = {
            "A13": "CONSIGNEE :",
            "B13": {"default": "DefaultCorp", "standard": "StdCorp", "daf": "DafCorp"},
        }
        p = TemplateClientProfileParser(hc)
        self.assertEqual(p.get_client_fullname(), "DefaultCorp")

    def test_empty_header_content(self):
        p = TemplateClientProfileParser({})
        self.assertIsNone(p.get_client_fullname())


class TestGetClientAddress(unittest.TestCase):

    def _parser(self):
        hc = {
            "A13": "CONSIGNEE :",
            "B13": "Client Name Ltd",
            "B14": "123 Industrial Park, Ho Chi Minh City, Vietnam.",
            "B15": "District 9, Block B",
            "A16": "SHIP: ",
            "B16": "BY TRUCK FROM BAVET",
        }
        return TemplateClientProfileParser(hc)

    def test_returns_address_rows_joined_with_newline(self):
        result = self._parser().get_client_address()
        self.assertIn("123 Industrial Park", result)
        self.assertIn("District 9", result)
        self.assertIn("\n", result)

    def test_excludes_consignee_name_row(self):
        result = self._parser().get_client_address()
        self.assertNotIn("Client Name Ltd", result)

    def test_excludes_ship_row_and_below(self):
        result = self._parser().get_client_address()
        self.assertNotIn("BY TRUCK FROM BAVET", result)

    def test_missing_consignee_returns_none(self):
        p = TemplateClientProfileParser({"A1": "INVOICE"})
        self.assertIsNone(p.get_client_address())

    def test_no_rows_between_consignee_and_ship_returns_none(self):
        hc = {
            "A13": "CONSIGNEE :",
            "B13": "Name Only",
            "A14": "SHIP:",
            "B14": "BY TRUCK",
        }
        p = TemplateClientProfileParser(hc)
        self.assertIsNone(p.get_client_address())

    def test_contact_lines_excluded_from_address(self):
        hc = {
            "A13": "CONSIGNEE :",
            "B13": "SomeCorp",
            "B14": "Lot 5 Industrial Zone, Vietnam",
            "B15": "P+84 650 3655 200 EXT:7074 MS EMILY",
            "A16": "SHIP:",
        }
        p = TemplateClientProfileParser(hc)
        addr = p.get_client_address()
        self.assertIn("Lot 5 Industrial Zone", addr)
        self.assertNotIn("MS EMILY", addr)


class TestGetClientContact(unittest.TestCase):

    def _hc(self):
        return {
            "A13": "CONSIGNEE :",
            "B13": "SomeCorp",
            "B14": "123 Street, City",
            "B15": "P+84 650 3655 200 EXT:7074 MS EMILY",
            "B16": "email@client.com",
            "A17": "SHIP:",
        }

    def test_phone_line_classified_as_contact(self):
        p = TemplateClientProfileParser(self._hc())
        contact = p.get_client_contact()
        self.assertIn("P+84", contact)

    def test_email_classified_as_contact(self):
        p = TemplateClientProfileParser(self._hc())
        self.assertIn("email@client.com", p.get_client_contact())

    def test_address_line_not_in_contact(self):
        p = TemplateClientProfileParser(self._hc())
        self.assertNotIn("123 Street", p.get_client_contact())

    def test_no_contact_lines_returns_none(self):
        hc = {
            "A13": "CONSIGNEE :",
            "B13": "Corp",
            "B14": "123 Street",
            "A15": "SHIP:",
        }
        p = TemplateClientProfileParser(hc)
        self.assertIsNone(p.get_client_contact())

    def test_tel_prefix_classified_as_contact(self):
        hc = {
            "A13": "CONSIGNEE :",
            "B13": "Corp",
            "B14": "Tel: 0274 3803833",
            "A15": "SHIP:",
        }
        p = TemplateClientProfileParser(hc)
        self.assertIsNotNone(p.get_client_contact())
        self.assertIn("Tel:", p.get_client_contact())


class TestGetShippingMethod(unittest.TestCase):

    def test_label_and_value_separate_cells(self):
        hc = {"A16": "SHIP: ", "B16": "BY TRUCK FROM BAVET TO HO CHI MINH"}
        p = TemplateClientProfileParser(hc)
        self.assertEqual(p.get_shipping_method(), "BY TRUCK FROM BAVET TO HO CHI MINH")

    def test_inline_in_same_cell(self):
        hc = {"A16": "SHIP: BY TRUCK FROM BAVET"}
        p = TemplateClientProfileParser(hc)
        self.assertIn("BY TRUCK", p.get_shipping_method())

    def test_missing_ship_label_returns_none(self):
        hc = {"A1": "INVOICE", "B1": "No 123"}
        p = TemplateClientProfileParser(hc)
        self.assertIsNone(p.get_shipping_method())

    def test_shipping_keyword_variant(self):
        hc = {"A16": "SHIPPING", "B16": "BY SEA"}
        p = TemplateClientProfileParser(hc)
        self.assertEqual(p.get_shipping_method(), "BY SEA")


# ---------------------------------------------------------------------------
# Integration smoke tests — real template JSON files
# ---------------------------------------------------------------------------

class TestRealTemplateFiles(unittest.TestCase):
    """
    Smoke-test: parser must not crash on any bundled *_template.json.
    Checks that all return values are either str or None (never raises).
    """

    def _run_one(self, tpl_path: Path):
        header_content = _load_header_content(tpl_path)
        label = f"{tpl_path.parent.name}/{tpl_path.name}"

        if not header_content:
            self.skipTest(f"{label}: no header_content in Invoice sheet")

        p = TemplateClientProfileParser(header_content)

        fullname = p.get_client_fullname()
        address  = p.get_client_address()
        contact  = p.get_client_contact()
        shipping = p.get_shipping_method()

        # Type assertions — each field must be str or None
        for field_name, value in [
            ("fullname", fullname),
            ("address",  address),
            ("contact",  contact),
            ("shipping", shipping),
        ]:
            self.assertIsInstance(
                value, (str, type(None)),
                msg=f"{label}: {field_name} returned unexpected type {type(value)}"
            )


def _make_test_method(path: Path):
    """Dynamically create a test method for each template file."""
    def test_method(self):
        self._run_one(path)
    test_method.__name__ = f"test_{path.parent.name}_{path.stem}"
    return test_method


# Dynamically attach one test per template file to TestRealTemplateFiles
for _tpl in _all_template_files():
    _method = _make_test_method(_tpl)
    setattr(TestRealTemplateFiles, _method.__name__, _method)


if __name__ == "__main__":
    unittest.main(verbosity=2)
