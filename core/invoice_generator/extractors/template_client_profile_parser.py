"""
template_client_profile_parser.py

Parses the Invoice sheet ``header_content`` dict from a ``*_template.json``
file to extract client profile data:

    - Client full name   (CONSIGNEE name row)
    - Client address     (street / location rows only)
    - Client contact     (phone / email / person rows)
    - Shipping method    (SHIP row)

The caller is responsible for passing only the ``header_content`` dict,
i.e. ``template_json["template_layout"]["Invoice"]["header_content"]``.

Usage::

    from core.invoice_generator.extractors.template_client_profile_parser import TemplateClientProfileParser

    with open("CLW_KH_template.json", encoding="utf-8") as f:
        template_json = json.load(f)

    header_content = template_json["template_layout"]["Invoice"]["header_content"]
    parser = TemplateClientProfileParser(header_content)
    print(parser.get_client_fullname())   # "Wanek Furniture Co., LTD."
    print(parser.get_client_address())    # "Lot D_5A_CN... Ho Chi Minh City."
    print(parser.get_client_contact())    # "P+84 650 3655 200 EXT:7074 MS EMILY..."
    print(parser.get_shipping_method())   # "BY TRUCK FROM BAVET..."
"""

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class TemplateClientProfileParser:
    """
    Extracts client profile fields from an Invoice ``header_content`` dict.

    The caller extracts the relevant section from ``*_template.json`` and
    passes it directly::

        header_content = template_json["template_layout"]["Invoice"]["header_content"]
        parser = TemplateClientProfileParser(header_content)

    Args:
        header_content: The ``header_content`` (or ``template_header_content``)
                        dict from the Invoice sheet of a ``*_template.json`` file.
                        Keys are Excel cell addresses (e.g. ``"B13"``), values are
                        plain strings or override dicts.
    """

    # Labels used to locate the CONSIGNEE block
    _CONSIGNEE_LABELS: tuple[str, ...] = ("CONSIGNEE", "CONSIGNEE :")
    # Labels used to locate the SHIP / shipping-method row
    _SHIP_LABELS: tuple[str, ...] = ("SHIP", "SHIP:", "SHIP: ", "SHIPPING")

    def __init__(self, header_content: Dict[str, Any]) -> None:
        self._header_content: Dict[str, Any] = header_content or {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_client_fullname(self) -> Optional[str]:
        """
        Return the consignee / client company name.

        Heuristic: find the cell whose **column** is immediately to the right
        of the ``CONSIGNEE`` label cell on the same row.  If the label and the
        value are merged into one cell (e.g. ``"Consignee : Wanek..."``), the
        part after the colon is returned.

        Returns:
            A single string with the client name, or ``None`` if not found.
        """
        label_cell, label_row = self._find_label_cell(self._CONSIGNEE_LABELS)
        if label_cell is None:
            logger.debug("[TemplateClientProfileParser] CONSIGNEE label not found.")
            return None

        # Case 1 – inline value in the same cell ("Consignee : Wanek Co.")
        inline = self._extract_inline_value(label_cell)
        if inline:
            return inline

        # Case 2 – value is in the sibling cell on the same row
        sibling = self._get_row_sibling_value(label_row, label_cell)
        if sibling:
            return sibling

        logger.debug("[TemplateClientProfileParser] CONSIGNEE name cell not found.")
        return None

    def get_client_address(self) -> Optional[str]:
        """
        Return only the street / location address lines as a single string.

        Lines that look like contact info (phone numbers, email addresses,
        person names with ``EXT:`` / ``MS `` / ``MR `` etc.) are excluded.
        Use :meth:`get_client_contact` to retrieve those.

        Returns:
            A newline-joined address string, or ``None`` if nothing is found.
        """
        address_lines, _ = self._classify_consignee_rows()
        return "\n".join(address_lines) if address_lines else None

    def get_client_contact(self) -> Optional[str]:
        """
        Return the contact info lines (phone, email, person names) as a single string.

        Lines are collected from the CONSIGNEE block and classified as contact
        info when they match any of the following heuristics:

        - Starts with a phone prefix (``P+``, ``Tel:``, ``+855``, ``+84``, …)
        - Contains an ``@`` sign (email address)
        - Contains ``EXT:``, ``MS ``, ``MR ``, ``MRS `` (person / extension)

        Returns:
            A newline-joined contact string, or ``None`` if nothing is found.
        """
        _, contact_lines = self._classify_consignee_rows()
        return "\n".join(contact_lines) if contact_lines else None

    def get_shipping_method(self) -> Optional[str]:
        """
        Return the shipping method / route description.

        Looks for the cell adjacent to the ``SHIP`` label row and returns
        its value.

        Returns:
            A single string describing the shipping method, or ``None``.
        """
        label_cell, label_row = self._find_label_cell(self._SHIP_LABELS)
        if label_cell is None:
            logger.debug("[TemplateClientProfileParser] SHIP label not found.")
            return None

        # Case 1 – inline value
        inline = self._extract_inline_value(label_cell)
        if inline:
            return inline

        # Case 2 – sibling cell
        sibling = self._get_row_sibling_value(label_row, label_cell)
        if sibling:
            return sibling

        logger.debug("[TemplateClientProfileParser] SHIP value cell not found.")
        return None

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _classify_consignee_rows(self) -> Tuple[List[str], List[str]]:
        """
        Collect all rows between CONSIGNEE and SHIP, then split them into
        ``(address_lines, contact_lines)`` using :meth:`_is_contact_line`.

        Both :meth:`get_client_address` and :meth:`get_client_contact` call
        this so the scan only runs once per call.

        Returns:
            ``(address_lines, contact_lines)`` — each a list of plain strings.
        """
        _, consignee_row = self._find_label_cell(self._CONSIGNEE_LABELS)
        _, ship_row = self._find_label_cell(self._SHIP_LABELS)

        if consignee_row is None:
            return [], []

        address_lines: List[str] = []
        contact_lines: List[str] = []

        rows_by_number = self._group_cells_by_row()
        for row_num in sorted(rows_by_number.keys()):
            if row_num <= consignee_row:
                continue
            if ship_row is not None and row_num >= ship_row:
                break

            for _addr, cell_val in rows_by_number[row_num].items():
                cleaned = self._clean_value(cell_val)
                if not cleaned:
                    continue
                if self._is_label(cleaned, self._CONSIGNEE_LABELS + self._SHIP_LABELS):
                    continue

                if self._is_contact_line(cleaned):
                    contact_lines.append(cleaned)
                else:
                    address_lines.append(cleaned)

        return address_lines, contact_lines

    @staticmethod
    def _is_contact_line(text: str) -> bool:
        """
        Return ``True`` when *text* looks like contact info rather than a
        street address.

        Heuristics (any match → contact):

        * Phone prefix at the start: ``P+``, ``Tel``, ``TEL``, ``+`` followed
          by digits, or a standalone country-code prefix (``+855``, ``+84``).
        * Contains ``@`` → email address.
        * Contains ``EXT:`` → telephone extension.
        * Contains ``MS ``, ``MR ``, ``MRS `` → contact person salutation.
        * Starts with ``CONTACT PERSON`` or ``CONTACT:`` → contact label.
        * Contains ``FAX`` → fax number.
        * Starts with ``TAX CODE`` or ``TAX :`` → tax/registration code line.
        """
        upper = text.upper()

        # Phone-number patterns
        phone_patterns = [
            r"^P\+",           # P+84 650...
            r"^TEL[:\s]",      # Tel: / TEL:
            r"^\+\d{1,3}",    # +855 / +84 ...
        ]
        for pat in phone_patterns:
            if re.search(pat, text, re.IGNORECASE):
                return True

        # Email
        if "@" in text:
            return True

        # Extension / person markers
        contact_keywords = ("EXT:", "MS ", "MR ", "MRS ")
        if any(kw in upper for kw in contact_keywords):
            return True

        # Contact person / contact label lines  (JF, KB, MOTO)
        if re.search(r"^CONTACT\s*(PERSON)?\s*[:\-]?", upper):
            return True

        # Fax / tax-code lines  (YNZX)
        if "FAX" in upper:
            return True
        if re.search(r"^TAX\s*(CODE)?\s*[:\-]?", upper):
            return True

        return False

    def _group_cells_by_row(self) -> Dict[int, Dict[str, Any]]:
        """
        Return a ``{row_number: {cell_addr: raw_value}}`` mapping built from
        ``self._header_content``.

        Row number is parsed from the numeric part of the cell address
        (e.g. ``"B13"`` → row 13).
        """
        from openpyxl.utils import coordinate_to_tuple  # lightweight, no workbook needed

        rows: Dict[int, Dict[str, Any]] = {}
        for addr, val in self._header_content.items():
            try:
                row_idx, _ = coordinate_to_tuple(addr)
            except Exception:
                continue
            rows.setdefault(row_idx, {})[addr] = val
        return rows

    def _find_label_cell(
        self, labels: tuple[str, ...]
    ) -> tuple[Optional[Any], Optional[int]]:
        """
        Search ``header_content`` for a cell whose value matches one of *labels*.

        Returns:
            ``(raw_cell_value, row_number)`` of the label cell, or
            ``(None, None)`` when not found.
        """
        from openpyxl.utils import coordinate_to_tuple

        for addr, val in self._header_content.items():
            text = self._clean_value(val)
            if self._is_label(text, labels):
                try:
                    row_idx, _ = coordinate_to_tuple(addr)
                    return val, row_idx
                except Exception:
                    continue

        return None, None

    def _get_row_sibling_value(
        self, row_num: int, exclude_cell_val: Any
    ) -> Optional[str]:
        """
        Return the first non-label, non-empty value on *row_num* that is not
        *exclude_cell_val*.
        """
        rows_by_number = self._group_cells_by_row()
        row_cells = rows_by_number.get(row_num, {})

        for _addr, val in row_cells.items():
            if val is exclude_cell_val:
                continue
            cleaned = self._clean_value(val)
            if cleaned and not self._is_label(cleaned, self._CONSIGNEE_LABELS + self._SHIP_LABELS):
                return cleaned

        return None

    @staticmethod
    def _clean_value(val: Any) -> str:
        """
        Normalise a cell value to a plain string.

        Cell values may be plain strings, numbers, or override dicts of the
        form ``{"default": "...", "standard": "...", "daf": "..."}``.
        The ``"default"`` key is used when a dict is encountered.
        """
        if val is None:
            return ""
        if isinstance(val, dict):
            val = val.get("default", "")
        return str(val).strip()

    @staticmethod
    def _extract_inline_value(cell_val: Any) -> Optional[str]:
        """
        If *cell_val* contains a colon (``:``) and non-trivial content after
        it, return the part after the colon.  Returns ``None`` otherwise.

        Example: ``"Consignee : Wanek Co."`` → ``"Wanek Co."``
        """
        text = TemplateClientProfileParser._clean_value(cell_val)
        if ":" in text:
            parts = text.split(":", 1)
            after = parts[1].strip()
            if len(after) > 3:  # guard against labels like "CONSIGNEE :"
                return after
        return None

    @staticmethod
    def _is_label(text: str, labels: tuple[str, ...]) -> bool:
        """Return ``True`` if *text* (case-insensitive) matches any of *labels*."""
        upper = text.upper()
        return any(upper == lbl.upper() or upper.startswith(lbl.upper()) for lbl in labels)
