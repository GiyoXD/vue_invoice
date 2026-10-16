import logging
import decimal
from typing import List, Dict, Any, Tuple, Optional, Callable
from .engine import Aggregator, coerce_input_row
from .schema import (
    COLUMN_REDUCERS,
    DECIMAL_COLUMNS,
    INTEGER_COLUMNS,
    LEATHER_SUMMARY_COLUMNS,
    LEATHER_TYPES,
    STANDARD_REQUIRED_COLUMNS,
    CUSTOM_REQUIRED_COLUMNS,
    MISSING_PO_SENTINEL,
    MISSING_ITEM_SENTINEL,
)

__all__ = [
    "quantize_record",
    "finalize_standard_aggregation",
    "finalize_item_merged_aggregation",
    "aggregate_standard_by_po_item_price",
    "aggregate_custom_by_po_item",
    "aggregate_by_item_merge_po",
    "calculate_leather_summary",
    "aggregate_per_po_with_pallets",
    "execute_strategy",
]


def quantize_record(data_dict: Dict[str, Any], decimals_map: Optional[Dict[str, int]] = None) -> None:
    """
    Quantizes any Decimal field in data_dict in place using ROUND_HALF_UP.
    Decimals are rounded according to decimals_map (defaults to 2 for any unspecified column).
    """
    if decimals_map is None:
        decimals_map = {}
    for col_id in DECIMAL_COLUMNS:
        val = data_dict.get(col_id)
        if val is not None:
            try:
                if not isinstance(val, decimal.Decimal):
                    clean_s = str(val).replace(",", "").strip()
                    if not clean_s:
                        continue
                    val = decimal.Decimal(clean_s)
                if val.is_nan() or val.is_infinite():
                    continue
                dec_places = decimals_map.get(col_id) if decimals_map else None
                if dec_places is None:
                    dec_places = 4 if col_id == "col_unit_price" else 2
                data_dict[col_id] = val.quantize(
                    decimal.Decimal(10) ** -dec_places,
                    rounding=decimal.ROUND_HALF_UP
                )
            except Exception:
                pass


def _extract_po_item_key(
    row: Dict[str, Any],
    include_price: bool = True
) -> Tuple[Any, Any, Any, Optional[str]]:
    """Extracts a standardized (PO, Item, Price, Description) key tuple from a row."""
    po_raw, item_raw, desc_raw = row.get('col_po'), row.get('col_item'), row.get('col_desc')
    po_s = str(po_raw).strip() if po_raw is not None else ""
    item_s = str(item_raw).strip() if item_raw is not None else ""
    desc_s = str(desc_raw).strip() if desc_raw is not None else ""
    po_key = po_s if po_s else MISSING_PO_SENTINEL
    item_key = item_s if item_s else MISSING_ITEM_SENTINEL
    description_key = desc_s if desc_s else None

    price_val = row.get('col_unit_price') if include_price else None
    return (po_key, item_key, price_val, description_key)


def _item_key_selector(row: Dict[str, Any]) -> Tuple[Any]:
    item_raw = row.get('col_item')
    item_s = str(item_raw).strip() if item_raw is not None else ""
    return (item_s if item_s else MISSING_ITEM_SENTINEL,)


def _leather_key_selector(row: Dict[str, Any]) -> str:
    desc = str(row.get('col_desc', "")).upper() if row.get('col_desc') else ""
    return LEATHER_TYPES[0] if LEATHER_TYPES[0] in desc else LEATHER_TYPES[1]


def finalize_standard_aggregation(
    global_aggregation_map: Dict[Tuple[Any, Any, Any, Optional[str]], Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None,
) -> Dict[Tuple[Any, Any, Any, Optional[str]], Dict[str, Any]]:
    """
    Finalizes standard aggregation by recalculating col_amount = sqft * price
    when price_dec is not None, and quantizing records.
    """
    prefix = "[finalize_standard_aggregation]"
    if decimals_map is None:
        decimals_map = {"col_amount": 2, "col_qty_sf": 2, "col_net": 2, "col_cbm": 2}

    for (po, item, price_dec, desc), val_dict in global_aggregation_map.items():
        if price_dec is not None:
            try:
                price = price_dec if isinstance(price_dec, decimal.Decimal) else decimal.Decimal(str(price_dec).replace(",", "").strip())
                sqft_val = val_dict.get('col_qty_sf')
                if isinstance(sqft_val, decimal.Decimal):
                    sqft = sqft_val
                elif sqft_val is not None and str(sqft_val).strip() != "":
                    sqft = decimal.Decimal(str(sqft_val).replace(",", "").strip())
                else:
                    sqft = decimal.Decimal(0)
                val_dict['col_amount'] = sqft * price
            except Exception as e:
                logging.warning(f"{prefix} Failed to calculate amount for key ({po}, {item}, {price_dec}, {desc}): {e}")
        quantize_record(val_dict, decimals_map)

    return global_aggregation_map


def aggregate_standard_by_po_item_price(
    processed_data: List[Dict[str, Any]],
    global_aggregation_map: Dict[Tuple[Any, Any, Any, Optional[str]], Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None,
    finalize: bool = True,
) -> Dict[Tuple[Any, Any, Any, Optional[str]], Dict[str, Any]]:
    """
    STANDARD Aggregation: Aggregates 'sqft' AND 'amount' values based on unique
    combinations of 'po', 'item', 'unit' price, AND 'description'.
    Updates the global_aggregation_map in place.
    """
    prefix = "[aggregate_standard]"
    required_cols = STANDARD_REQUIRED_COLUMNS

    if decimals_map is None:
        decimals_map = {"col_amount": 2, "col_qty_sf": 2, "col_net": 2, "col_cbm": 2}

    if not processed_data:
        logging.info(f"{prefix} No data rows found in this table. Global map unchanged.")
        return global_aggregation_map

    # Check for required columns existing in at least one row
    missing_cols = [col for col in required_cols if not any(col in row for row in processed_data)]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform STANDARD aggregation: Missing required columns {missing_cols}. Skipping this table.")
        return global_aggregation_map

    aggregator = Aggregator(key_selector=_extract_po_item_key, reducers=COLUMN_REDUCERS)
    aggregator.aggregate(processed_data, global_aggregation_map)

    if finalize:
        finalize_standard_aggregation(global_aggregation_map, decimals_map)

    logging.debug(f"{prefix} Finished processing rows.")
    return global_aggregation_map


def aggregate_custom_by_po_item(
    processed_data: List[Dict[str, Any]],
    global_custom_aggregation_map: Dict[Tuple[Any, Any, Any, Optional[str]], Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None
) -> Dict[Tuple[Any, Any, Any, Optional[str]], Dict[str, Any]]:
    """
    CUSTOM Aggregation: Aggregates 'sqft' and 'amount' values based on unique
    combinations of 'po', 'item', AND 'description'. Uses a 4-element key
    (PO, Item, None, Description) for structural consistency with standard aggregation.
    Updates the global_custom_aggregation_map in place.
    """
    prefix = "[aggregate_custom]"
    required_cols = CUSTOM_REQUIRED_COLUMNS

    if not processed_data:
        logging.info(f"{prefix} No data rows found in this table. Global custom aggregation map remains unchanged.")
        return global_custom_aggregation_map

    # Check for required columns existing in at least one row
    missing_cols = [col for col in required_cols if not any(col in row for row in processed_data)]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform full CUSTOM aggregation: Missing required columns {missing_cols}. Proceeding cautiously.")

    aggregator = Aggregator(
        key_selector=lambda row: _extract_po_item_key(row, include_price=False),
        reducers=COLUMN_REDUCERS,
    )
    aggregator.aggregate(processed_data, global_custom_aggregation_map)

    for val_dict in global_custom_aggregation_map.values():
        quantize_record(val_dict, decimals_map)

    logging.debug(f"{prefix} Finished processing rows.")
    return global_custom_aggregation_map


def finalize_item_merged_aggregation(
    global_item_aggregation_map: Dict[Tuple[Any], Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None,
) -> Dict[Tuple[Any], Dict[str, Any]]:
    """
    Finalizes item-merged aggregation by calculating unit price from amount and sqft,
    ensuring col_po is a string, and quantizing records.
    """
    for val_dict in global_item_aggregation_map.values():
        if 'col_po' not in val_dict or val_dict['col_po'] is None:
            val_dict['col_po'] = ""
        amt = val_dict.get('col_amount', decimal.Decimal(0))
        sf = val_dict.get('col_qty_sf', decimal.Decimal(0))
        try:
            if sf != 0:
                val_dict['col_unit_price'] = amt / sf
            else:
                val_dict['col_unit_price'] = decimal.Decimal(0)
        except Exception:
            val_dict['col_unit_price'] = decimal.Decimal(0)
        quantize_record(val_dict, decimals_map)
    return global_item_aggregation_map


def aggregate_by_item_merge_po(
    processed_data: List[Dict[str, Any]],
    global_item_aggregation_map: Dict[Tuple[Any], Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None,
    finalize: bool = True,
) -> Dict[Tuple[Any], Dict[str, Any]]:
    """
    ITEM_MERGED Aggregation: Aggregates values based on unique col_item,
    merging unique POs into a 'col_po' string joined by ' / '.
    Reducers:
      - col_qty_sf, col_amount, col_net, col_cbm: decimal_sum_reducer
      - col_qty_pcs, col_pallet_count: int_sum_reducer
      - col_dc, col_desc: first_non_empty_reducer
      - col_po: po_merge_reducer
    Updates the global_item_aggregation_map in place and returns it.
    """
    prefix = "[aggregate_by_item_merge_po]"
    if not processed_data:
        logging.info(f"{prefix} No data rows found in this table. Global item aggregation map remains unchanged.")
        return global_item_aggregation_map

    aggregator = Aggregator(key_selector=_item_key_selector, reducers=COLUMN_REDUCERS)
    aggregator.aggregate(processed_data, global_item_aggregation_map)

    if finalize:
        finalize_item_merged_aggregation(global_item_aggregation_map, decimals_map)

    logging.debug(f"{prefix} Finished processing rows.")
    return global_item_aggregation_map


def calculate_leather_summary(
    processed_data: List[Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None
) -> Dict[str, Any]:
    """
    Calculates the leather summary (PCS, SQFT, Net, Gross, Pallet Count) per leather type.
    Iterates through rows to sum values for each leather type found in 'description' or 'desc'.
    BUFFALO = rows containing "BUFFALO" in description
    COW = rows NOT containing "BUFFALO" (all other leather)
    """
    leather_cols = LEATHER_SUMMARY_COLUMNS
    summary = {
        k: {c: 0 if c in INTEGER_COLUMNS else decimal.Decimal(0) for c in leather_cols}
        for k in LEATHER_TYPES
    }

    if not processed_data:
        for l_data in summary.values():
            quantize_record(l_data, decimals_map)
        return summary

    aggregator = Aggregator(key_selector=_leather_key_selector, reducers={c: COLUMN_REDUCERS[c] for c in leather_cols})
    aggregator.aggregate(processed_data, summary)

    for l_data in summary.values():
        quantize_record(l_data, decimals_map)

    return summary


def aggregate_per_po_with_pallets(
    processed_data: List[Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None
) -> List[Dict[str, Any]]:
    """
    Aggregates data by PO, Item, Unit Price, and Description (same as STANDARD aggregation),
    summing pallet, pcs, sqft, amount, net, gross, cbm.
    """
    if not processed_data:
        return []

    from .serializer import format_aggregation_as_list

    aggregator = Aggregator(key_selector=_extract_po_item_key, reducers=COLUMN_REDUCERS)
    agg_map = aggregator.aggregate(processed_data)

    for val in agg_map.values():
        quantize_record(val, decimals_map)

    result = format_aggregation_as_list(agg_map)
    
    logging.info(f"[aggregate_per_po_with_pallets] Aggregated {len(processed_data)} rows into {len(result)} unique PO+Item+Price combinations.")
    return result


def execute_strategy(
    strategy: str,
    data: List[Dict[str, Any]],
    decimals_map: Optional[Dict[str, int]] = None,
    return_raw: bool = False,
    chunk_size: int = 2,
    item_separator: str = "/",
    line_separator: str = "\n",
) -> Any:
    """
    Executes a specified aggregation strategy on row data.

    Supported strategies:
        - 'standard': PO, Item, Unit Price, Description aggregation
        - 'custom': PO, Item, Description aggregation
        - 'item_merged': Item-merged aggregation (combining POs)
        - 'pallets' or 'manifest_by_pallet_per_po': Manifest by pallet per PO aggregation
        - 'leather_summary': Leather summary by leather type (BUFFALO vs COW)
        - 'daf': DAF Compounding aggregation
    """
    data = [coerce_input_row(r) for r in data if r]
    from .serializer import format_aggregation_as_list

    strategy_key = str(strategy).strip().lower()

    map_strategies: Dict[str, Callable[..., Any]] = {
        'standard': aggregate_standard_by_po_item_price,
        'custom': aggregate_custom_by_po_item,
        'item_merged': aggregate_by_item_merge_po,
    }
    if strategy_key in map_strategies:
        agg_map: Dict[Any, Dict[str, Any]] = {}
        map_strategies[strategy_key](data, agg_map, decimals_map=decimals_map)
        return agg_map if return_raw else format_aggregation_as_list(agg_map)

    if strategy_key in ('pallets', 'manifest_by_pallet_per_po'):
        return aggregate_per_po_with_pallets(data, decimals_map=decimals_map)

    if strategy_key == 'leather_summary':
        return calculate_leather_summary(data, decimals_map=decimals_map)

    if strategy_key == 'daf':
        from .daf import perform_DAF_compounding
        return perform_DAF_compounding(
            data,
            chunk_size=chunk_size,
            item_separator=item_separator,
            line_separator=line_separator,
            decimals_map=decimals_map,
        )

    raise ValueError(f"Unknown aggregation strategy: '{strategy}'")

