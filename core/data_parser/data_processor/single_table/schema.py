import decimal
from typing import Any, Callable, Dict, Tuple, TypedDict

from .reducers import (
    decimal_sum_reducer,
    first_non_empty_reducer,
    first_value_reducer,
    int_sum_reducer,
    unique_join_reducer,
)

MISSING_ITEM_SENTINEL: str = "<MISSING_ITEM>"
MISSING_PO_SENTINEL: str = "<MISSING_PO>"

DECIMAL_COLUMNS: Tuple[str, ...] = (
    "col_amount",
    "col_cbm",
    "col_gross",
    "col_net",
    "col_qty_sf",
    "col_unit_price",
)
INTEGER_COLUMNS: Tuple[str, ...] = ("col_qty_pcs", "col_pallet_count")
LEATHER_TYPES: Tuple[str, ...] = ("BUFFALO", "COW")
TEXT_COLUMNS: Tuple[str, ...] = ("col_dc", "col_desc", "col_item", "col_po")

CUSTOM_REQUIRED_COLUMNS: Tuple[str, ...] = ("col_po", "col_item", "col_qty_sf", "col_amount")
INVOICE_FOOTER_COLUMNS: Tuple[str, ...] = (
    "col_qty_pcs",
    "col_pallet_count",
    "col_qty_sf",
    "col_amount",
    "col_net",
    "col_gross",
    "col_cbm",
)
INVOICE_LINE_ITEM_COLUMNS: Tuple[str, ...] = (
    "col_po",
    "col_item",
    "col_unit_price",
    "col_desc",
    "col_qty_pcs",
    "col_qty_sf",
    "col_amount",
    "col_pallet_count",
    "col_net",
    "col_gross",
    "col_cbm",
    "col_dc",
)
LEATHER_SUMMARY_COLUMNS: Tuple[str, ...] = (
    "col_qty_pcs",
    "col_qty_sf",
    "col_net",
    "col_gross",
    "col_cbm",
    "col_pallet_count",
)
STANDARD_REQUIRED_COLUMNS: Tuple[str, ...] = (
    "col_po",
    "col_item",
    "col_unit_price",
    "col_qty_sf",
    "col_amount",
)

COLUMN_REDUCERS: Dict[str, Callable[..., Any]] = {
    "col_po": unique_join_reducer,
    **{col: first_non_empty_reducer for col in TEXT_COLUMNS if col != "col_po"},
    **{col: int_sum_reducer for col in INTEGER_COLUMNS},
    **{col: decimal_sum_reducer for col in DECIMAL_COLUMNS if col != "col_unit_price"},
    "col_unit_price": first_value_reducer,
}
DEFAULT_COLUMN_REDUCERS = COLUMN_REDUCERS

DAF_CHUNK_COLUMNS: Tuple[str, ...] = ("col_po", "col_item", "col_dc")
DAF_SUM_COLUMNS: Tuple[str, ...] = ("col_qty_sf", "col_amount", "col_net")


class DAFCompoundingResult(TypedDict):
    col_po: str
    col_item: str
    col_desc: str
    col_qty_sf: decimal.Decimal
    col_amount: decimal.Decimal
    col_net: decimal.Decimal
    col_dc: str
    material_type: str

__all__ = [
    "COLUMN_REDUCERS",
    "CUSTOM_REQUIRED_COLUMNS",
    "DAFCompoundingResult",
    "DAF_CHUNK_COLUMNS",
    "DAF_SUM_COLUMNS",
    "DECIMAL_COLUMNS",
    "DEFAULT_COLUMN_REDUCERS",
    "INTEGER_COLUMNS",
    "INVOICE_FOOTER_COLUMNS",
    "INVOICE_LINE_ITEM_COLUMNS",
    "LEATHER_SUMMARY_COLUMNS",
    "LEATHER_TYPES",
    "MISSING_ITEM_SENTINEL",
    "MISSING_PO_SENTINEL",
    "STANDARD_REQUIRED_COLUMNS",
    "TEXT_COLUMNS",
]
