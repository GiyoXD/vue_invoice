from collections.abc import Mapping, MutableMapping
from dataclasses import dataclass, fields
from decimal import Decimal, InvalidOperation
import math
import re
from typing import Any, Dict, Iterator, List, Optional, Tuple, Union


def _parse_str(val: Any) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return None
        try:
            if val.is_integer():
                return str(int(val))
        except (OverflowError, ValueError):
            pass
    if isinstance(val, Decimal):
        try:
            if val.is_nan() or val.is_infinite():
                return None
            if val == val.to_integral():
                return str(int(val))
        except (OverflowError, InvalidOperation, ValueError):
            return None
    try:
        s = str(val).strip()
        if not s or s.lower() in ("nan", "none", "inf", "-inf", "+inf", "infinity", "-infinity", "+infinity"):
            return None
        return s
    except (OverflowError, InvalidOperation, ValueError, TypeError):
        return None


def _parse_int(val: Any) -> Optional[int]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, int):
        return val
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return None
        try:
            return int(round(val))
        except (OverflowError, ValueError):
            return None
    if isinstance(val, Decimal):
        try:
            if val.is_nan() or val.is_infinite():
                return None
            return int(round(val))
        except (OverflowError, InvalidOperation, ValueError):
            return None
    if isinstance(val, str):
        s = val.strip().replace(",", "")
        if not s or s.lower() in ("nan", "none", "inf", "-inf", "+inf", "infinity", "-infinity", "+infinity"):
            return None
        try:
            f = float(s)
            if math.isnan(f) or math.isinf(f):
                return None
            return int(round(f))
        except (ValueError, TypeError, OverflowError):
            return None
    try:
        return int(val)
    except (ValueError, TypeError, OverflowError, InvalidOperation):
        return None


def _parse_decimal(val: Any) -> Optional[Decimal]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, Decimal):
        try:
            if val.is_nan() or val.is_infinite():
                return None
            return val
        except (InvalidOperation, OverflowError):
            return None
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return None
        try:
            return Decimal(str(val))
        except (InvalidOperation, OverflowError, ValueError):
            return None
    if isinstance(val, int):
        try:
            return Decimal(val)
        except (InvalidOperation, OverflowError, ValueError):
            return None
    if isinstance(val, str):
        s = val.strip().lstrip('=').strip().replace(",", "").replace("$", "")
        if not s or s.lower() in ("nan", "none", "-", "n/a", "inf", "-inf", "+inf", "infinity", "-infinity", "+infinity"):
            return None
        try:
            d = Decimal(s)
            if d.is_nan() or d.is_infinite():
                return None
            return d
        except (InvalidOperation, TypeError, ValueError, OverflowError):
            return None
    try:
        d = Decimal(str(val))
        if d.is_nan() or d.is_infinite():
            return None
        return d
    except (InvalidOperation, TypeError, ValueError, OverflowError, Exception):
        return None


def _parse_pallet_count(val: Any) -> Optional[Union[int, str]]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, int):
        return val
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return None
        try:
            return int(round(val))
        except (OverflowError, ValueError):
            return None
    if isinstance(val, Decimal):
        try:
            if val.is_nan() or val.is_infinite():
                return None
            return int(round(val))
        except (OverflowError, InvalidOperation, ValueError):
            return None
    if isinstance(val, str):
        s = val.strip()
        if not s or s.lower() in ("nan", "none", "n/a", "inf", "-inf", "+inf", "infinity", "-infinity", "+infinity"):
            return None
        # If string is pure digits, parse to int; if range like "1-2", preserve string!
        if s.isdigit():
            try:
                return int(s)
            except (OverflowError, ValueError):
                return s
        try:
            f = float(s)
            if not math.isnan(f) and not math.isinf(f) and f.is_integer() and ("." in s and not any(c in s for c in "-~")):
                return int(f)
        except (ValueError, OverflowError):
            pass
        return s
    return _parse_str(val)


def _parse_col_no(val: Any) -> Optional[Union[int, str]]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, int):
        return val
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return None
        try:
            if val.is_integer():
                return int(val)
        except (OverflowError, ValueError):
            pass
        return str(val)
    if isinstance(val, Decimal):
        try:
            if val.is_nan() or val.is_infinite():
                return None
            if val == val.to_integral():
                return int(val)
        except (OverflowError, InvalidOperation, ValueError):
            pass
        return str(val)
    if isinstance(val, str):
        s = val.strip()
        if not s or s.lower() in ("nan", "none", "n/a", "inf", "-inf", "+inf", "infinity", "-infinity", "+infinity"):
            return None
        # Preserve strings like "001", "01" with leading zeros
        if s.isdigit():
            if len(s) > 1 and s.startswith("0"):
                return s
            try:
                return int(s)
            except (OverflowError, ValueError):
                return s
        return s
    return _parse_str(val)


def _eval_cbm_formula(s: Any, row_idx: int = 0) -> Tuple[Optional[Decimal], Optional[str]]:
    from core.data_parser.data_processor.cbm import _eval_cbm_formula as _eval_impl
    return _eval_impl(s, row_idx)


_FIELD_NAMES: Tuple[str, ...] = (
    "col_po",
    "col_item",
    "col_production_order_no",
    "col_batch_no",
    "col_line_no",
    "col_reference_code",
    "col_manual_no",
    "col_no",
    "col_desc",
    "col_dc",
    "col_level",
    "col_grade",
    "col_direction",
    "col_static",
    "col_remarks",
    "col_qty_pcs",
    "col_pallet_count",
    "col_pallet_no",
    "col_pallet_id",
    "col_qty_header",
    "col_qty_sf",
    "col_unit_sf",
    "col_sqm",
    "col_net",
    "col_gross",
    "col_net_raw",
    "col_gross_raw",
    "col_cbm",
    "col_cbm_raw",
    "col_unit_price",
    "col_amount",
    "col_container_no",
    "col_hs_code",
    "col_inv_no",
    "col_inv_date",
    "col_inv_ref",
    "col_production_date",
    "col_date_recipt",
    "_row_num",
)
_FIELD_SET = set(_FIELD_NAMES)

_DECIMAL_FIELDS = {
    "col_qty_sf",
    "col_unit_sf",
    "col_sqm",
    "col_net",
    "col_gross",
    "col_net_raw",
    "col_gross_raw",
    "col_cbm",
    "col_unit_price",
    "col_amount",
}

_INT_FIELDS = {
    "col_qty_pcs",
    "col_qty_header",
    "_row_num",
}

_PARSERS: Dict[str, Any] = {
    "col_no": _parse_col_no,
    "col_pallet_count": _parse_pallet_count,
    **{f: _parse_decimal for f in _DECIMAL_FIELDS},
    **{f: _parse_int for f in _INT_FIELDS},
}


@dataclass(slots=True)
class InvoiceRow:
    # Identifiers:
    col_po: Optional[str] = None
    col_item: Optional[str] = None
    col_production_order_no: Optional[str] = None
    col_batch_no: Optional[str] = None
    col_line_no: Optional[str] = None
    col_reference_code: Optional[str] = None
    col_manual_no: Optional[str] = None
    col_no: Optional[Union[int, str]] = None

    # Metadata & Descriptions:
    col_desc: Optional[str] = None
    col_dc: Optional[str] = None
    col_level: Optional[str] = None
    col_grade: Optional[str] = None
    col_direction: Optional[str] = None
    col_static: Optional[str] = None
    col_remarks: Optional[str] = None

    # Counts:
    col_qty_pcs: Optional[int] = None
    col_pallet_count: Optional[Union[int, str]] = None
    col_pallet_no: Optional[str] = None
    col_pallet_id: Optional[str] = None
    col_qty_header: Optional[int] = None

    # Measurements & Weights (Decimal):
    col_qty_sf: Optional[Decimal] = None
    col_unit_sf: Optional[Decimal] = None
    col_sqm: Optional[Decimal] = None
    col_net: Optional[Decimal] = None
    col_gross: Optional[Decimal] = None
    col_net_raw: Optional[Decimal] = None
    col_gross_raw: Optional[Decimal] = None
    col_cbm: Optional[Decimal] = None
    col_cbm_raw: Optional[str] = None

    # Financials (Decimal):
    col_unit_price: Optional[Decimal] = None
    col_amount: Optional[Decimal] = None

    # Logistics & Customs:
    col_container_no: Optional[str] = None
    col_hs_code: Optional[str] = None
    col_inv_no: Optional[str] = None
    col_inv_date: Optional[str] = None
    col_inv_ref: Optional[str] = None
    col_production_date: Optional[str] = None
    col_date_recipt: Optional[str] = None

    # Internal metadata:
    _row_num: Optional[int] = None

    @classmethod
    def from_raw_dict(cls, data: Dict[str, Any]) -> "InvoiceRow":
        """Parses raw boundary dict, coerces strings, ints, Decimals safely once at boundary, returns InvoiceRow."""
        row_num_val = data.get("_row_num") if "_row_num" in data else data.get("row_num")
        try:
            row_idx = int(row_num_val or 0)
        except (ValueError, TypeError):
            row_idx = 0
        raw_cbm = data.get("col_cbm")
        raw_cbm_raw = data.get("col_cbm_raw")

        cbm_raw_val = _parse_str(raw_cbm_raw)
        cbm_f_val, cbm_f_str = _eval_cbm_formula(raw_cbm, row_idx)

        if cbm_f_val is not None:
            cbm_val = cbm_f_val
            if cbm_raw_val is None:
                cbm_raw_val = cbm_f_str
        else:
            cbm_val = _parse_decimal(raw_cbm)
            if cbm_val is None and cbm_raw_val is not None:
                rf_val, rf_str = _eval_cbm_formula(cbm_raw_val, row_idx)
                cbm_val = rf_val if rf_val is not None else _parse_decimal(cbm_raw_val)

        return cls(
            col_po=_parse_str(data.get("col_po")),
            col_item=_parse_str(data.get("col_item")),
            col_production_order_no=_parse_str(data.get("col_production_order_no")),
            col_batch_no=_parse_str(data.get("col_batch_no")),
            col_line_no=_parse_str(data.get("col_line_no")),
            col_reference_code=_parse_str(data.get("col_reference_code")),
            col_manual_no=_parse_str(data.get("col_manual_no")),
            col_no=_parse_col_no(data.get("col_no")),
            col_desc=_parse_str(data.get("col_desc")),
            col_dc=_parse_str(data.get("col_dc")),
            col_level=_parse_str(data.get("col_level")),
            col_grade=_parse_str(data.get("col_grade")),
            col_direction=_parse_str(data.get("col_direction")),
            col_static=_parse_str(data.get("col_static")),
            col_remarks=_parse_str(data.get("col_remarks")),
            col_qty_pcs=_parse_int(data.get("col_qty_pcs")),
            col_pallet_count=_parse_pallet_count(data.get("col_pallet_count")),
            col_pallet_no=_parse_str(data.get("col_pallet_no")),
            col_pallet_id=_parse_str(data.get("col_pallet_id")),
            col_qty_header=_parse_int(data.get("col_qty_header")),
            col_qty_sf=_parse_decimal(data.get("col_qty_sf")),
            col_unit_sf=_parse_decimal(data.get("col_unit_sf")),
            col_sqm=_parse_decimal(data.get("col_sqm")),
            col_net=_parse_decimal(data.get("col_net")),
            col_gross=_parse_decimal(data.get("col_gross")),
            col_net_raw=_parse_decimal(data.get("col_net_raw")),
            col_gross_raw=_parse_decimal(data.get("col_gross_raw")),
            col_cbm=cbm_val,
            col_cbm_raw=cbm_raw_val,
            col_unit_price=_parse_decimal(data.get("col_unit_price")),
            col_amount=_parse_decimal(data.get("col_amount")),
            col_container_no=_parse_str(data.get("col_container_no")),
            col_hs_code=_parse_str(data.get("col_hs_code")),
            col_inv_no=_parse_str(data.get("col_inv_no")),
            col_inv_date=_parse_str(data.get("col_inv_date")),
            col_inv_ref=_parse_str(data.get("col_inv_ref")),
            col_production_date=_parse_str(data.get("col_production_date")),
            col_date_recipt=_parse_str(data.get("col_date_recipt")),
            _row_num=_parse_int(row_num_val),
        )

    def to_dict(self) -> Dict[str, Any]:
        """Returns dictionary representation of all set attributes for serialization/export."""
        return {f: getattr(self, f) for f in _FIELD_NAMES if getattr(self, f) is not None}

    def __getitem__(self, key: str) -> Any:
        if key == "row_num":
            key = "_row_num"
        if key in _FIELD_SET:
            return getattr(self, key)
        raise KeyError(key)

    def __setitem__(self, key: str, value: Any) -> None:
        if key == "row_num":
            key = "_row_num"
        if key not in _FIELD_SET:
            raise KeyError(key)

        if key == "col_cbm":
            try:
                row_idx = int(self._row_num or 0)
            except (ValueError, TypeError):
                row_idx = 0
            cbm_val, cbm_raw_val = _eval_cbm_formula(value, row_idx)
            if cbm_val is not None:
                self.col_cbm = cbm_val
                self.col_cbm_raw = cbm_raw_val
            else:
                self.col_cbm = _parse_decimal(value)
            return
        elif key == "col_cbm_raw":
            self.col_cbm_raw = _parse_str(value)
            return

        parser = _PARSERS.get(key, _parse_str)
        parsed_val = parser(value)
        setattr(self, key, parsed_val)

    def get(self, key: str, default: Any = None) -> Any:
        if key == "row_num":
            key = "_row_num"
        if key in _FIELD_SET:
            val = getattr(self, key)
            return default if val is None else val
        return default

    def __iter__(self) -> Iterator[str]:
        return iter([f for f in _FIELD_NAMES if getattr(self, f) is not None])

    def __len__(self) -> int:
        return len([f for f in _FIELD_NAMES if getattr(self, f) is not None])

    def __contains__(self, key: Any) -> bool:
        if not isinstance(key, str):
            return False
        if key == "row_num":
            key = "_row_num"
        return key in _FIELD_SET and getattr(self, key) is not None

    def keys(self) -> List[str]:
        return [f for f in _FIELD_NAMES if getattr(self, f) is not None]

    def values(self) -> List[Any]:
        return [getattr(self, f) for f in _FIELD_NAMES if getattr(self, f) is not None]

    def __delitem__(self, key: str) -> None:
        if key == "row_num":
            key = "_row_num"
        if key in _FIELD_SET:
            setattr(self, key, None)
        else:
            raise KeyError(key)

    def items(self) -> List[Tuple[str, Any]]:
        return [(f, getattr(self, f)) for f in _FIELD_NAMES if getattr(self, f) is not None]


Mapping.register(InvoiceRow)
MutableMapping.register(InvoiceRow)
