import logging
import decimal
from typing import List, Dict, Any, Optional, Tuple, cast
from .aggregations import quantize_record
from .reducers import decimal_sum_reducer
from .schema import (
    DAFCompoundingResult,
    DAF_CHUNK_COLUMNS,
    DAF_SUM_COLUMNS,
    MISSING_PO_SENTINEL,
    MISSING_ITEM_SENTINEL,
)

FinalDAFResultType = List[DAFCompoundingResult]

__all__ = ["perform_DAF_compounding", "DAFCompoundingResult", "FinalDAFResultType"]


def format_chunks(items: List[str], size: int, item_sep: str, line_sep: str) -> str:
    if not items:
        return ""
    if size <= 0:
        raise ValueError("size must be > 0")
    return line_sep.join(
        item_sep.join(str(it) for it in items[i:i + size])
        for i in range(0, len(items), size)
    )


def _split_by_material(
    data: List[Dict[str, Any]],
) -> List[Tuple[str, List[Dict[str, Any]]]]:
    buffalo_rows: List[Dict[str, Any]] = []
    normal_rows: List[Dict[str, Any]] = []
    for r in data:
        if "BUFFALO" in str(r.get("col_desc") or "").upper():
            buffalo_rows.append(r)
        else:
            normal_rows.append(r)

    result: List[Tuple[str, List[Dict[str, Any]]]] = []
    if buffalo_rows:
        result.append(("buffalo", buffalo_rows))
    if normal_rows:
        result.append(("normal", normal_rows))
    return result


def _split_by_po_chunks(
    data: List[Dict[str, Any]],
) -> List[Tuple[str, List[Dict[str, Any]]]]:
    pos = sorted({(str(r.get("col_po")).strip() if r.get("col_po") is not None else "").strip() or MISSING_PO_SENTINEL for r in data})
    if len(pos) <= 7:
        return [("po_chunk", data)]

    mid = (len(pos) + 1) // 2
    first_half_pos = set(pos[:mid])
    chunk_0: List[Dict[str, Any]] = []
    chunk_1: List[Dict[str, Any]] = []
    for r in data:
        po_key = (str(r.get("col_po")).strip() if r.get("col_po") is not None else "").strip() or MISSING_PO_SENTINEL
        if po_key in first_half_pos:
            chunk_0.append(r)
        else:
            chunk_1.append(r)

    return [("po_chunk", chunk_0), ("po_chunk", chunk_1)]


def _create_empty_row(material_type: str = "normal") -> DAFCompoundingResult:
    row = {
        **{col: "" for col in DAF_CHUNK_COLUMNS},
        "col_desc": "",
        **{col: decimal.Decimal(0) for col in DAF_SUM_COLUMNS},
        "material_type": material_type,
    }
    return cast(DAFCompoundingResult, row)


def perform_DAF_compounding(
    data: List[Dict[str, Any]],
    chunk_size: int = 2,
    item_separator: str = "/",
    line_separator: str = "\n",
    decimals_map: Optional[Dict[str, int]] = None,
) -> FinalDAFResultType:
    prefix = "[perform_DAF_compounding]"
    logging.debug(f"{prefix} Starting DAF Compounding directly from row data.")

    try:
        chunk_size = int(chunk_size)
    except Exception:
        chunk_size = 2
    if chunk_size <= 0:
        chunk_size = 2

    if item_separator is None:
        item_separator = "/"
    if line_separator is None:
        line_separator = "\n"

    if not data:
        logging.warning(f"{prefix} Input data list is empty. Returning default empty DAF groups.")
        empty = _create_empty_row()
        quantize_record(cast(Dict[str, Any], empty), decimals_map)
        return [empty]

    # Business Note: When descriptions exist (has_desc is True), PO chunking is
    # intentionally ignored by design. Output splits strictly by material (buffalo vs normal/cow).
    has_desc = any(r.get("col_desc") and str(r.get("col_desc")).strip() for r in data)
    buckets = _split_by_material(data) if has_desc else _split_by_po_chunks(data)

    sentinel_map = {
        "col_po": MISSING_PO_SENTINEL,
        "col_item": MISSING_ITEM_SENTINEL,
    }

    results: FinalDAFResultType = []
    for mat_type, g_rows in buckets:
        if not g_rows:
            continue

        res: DAFCompoundingResult = _create_empty_row(mat_type)
        res_dict = cast(Dict[str, Any], res)

        for col in DAF_CHUNK_COLUMNS:
            sentinel = sentinel_map.get(col)
            exclusions = {"", sentinel} if sentinel else {""}
            col_items = sorted({str(r.get(col)).strip() for r in g_rows if r.get(col) is not None and str(r.get(col)).strip() != ""} - exclusions)
            res_dict[col] = format_chunks(col_items, chunk_size, item_separator, line_separator)

        if has_desc:
            descs = sorted({str(r.get("col_desc") or "").strip() for r in g_rows} - {""})
            res_dict["col_desc"] = format_chunks(descs, 1, "", "\n")

        for col in DAF_SUM_COLUMNS:
            res_dict[col] = decimal_sum_reducer([r.get(col) for r in g_rows])

        quantize_record(res_dict, decimals_map)
        results.append(res)

    logging.info(f"{prefix} DAF Compounding complete ({len(results)} chunks created).")
    return results
