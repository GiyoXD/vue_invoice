import logging
import decimal
from typing import List, Dict, Any, Tuple, Optional
from .engine import Aggregator
from .reducers import (
    int_sum_reducer,
    decimal_sum_reducer,
    first_non_empty_reducer,
)


def aggregate_standard_by_po_item_price(
    processed_data: List[Dict[str, Any]],
    global_aggregation_map: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]]
) -> Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]]:
    """
    STANDARD Aggregation: Aggregates 'sqft' AND 'amount' values based on unique
    combinations of 'po', 'item', 'unit' price, AND 'description'.
    Updates the global_aggregation_map in place.
    """
    prefix = "[aggregate_standard]"
    required_cols = ['col_po', 'col_item', 'col_unit_price', 'col_qty_sf', 'col_amount']

    if not processed_data:
        logging.info(f"{prefix} No data rows found in this table. Global map unchanged.")
        return global_aggregation_map

    # Check for required columns existing in at least one row
    missing_cols = [col for col in required_cols if not any(col in row for row in processed_data)]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform STANDARD aggregation: Missing required columns {missing_cols}. Skipping this table.")
        return global_aggregation_map

    # Key selector
    def standard_key_selector(row):
        po_val, item_val = row.get('col_po'), row.get('col_item')
        unit_price_raw = row.get('col_unit_price')
        desc_raw = row.get('col_desc')
        
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw
        description_key = description_key if description_key else None

        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"

        price_dec = unit_price_raw
        return (po_key, item_key, price_dec, description_key)

    # Reducers
    reducers = {
        'col_qty_sf': ('col_qty_sf', decimal_sum_reducer),
        'col_amount': ('col_amount', decimal_sum_reducer),
        'col_net': ('col_net', decimal_sum_reducer),
        'col_cbm': ('col_cbm', decimal_sum_reducer)
    }

    # Run Aggregator
    aggregator = Aggregator(key_selector=standard_key_selector, reducers=reducers)
    aggregator.aggregate(processed_data, global_aggregation_map)

    logging.info(f"{prefix} Finished processing rows.")
    return global_aggregation_map


def aggregate_custom_by_po_item(
    processed_data: List[Dict[str, Any]],
    global_custom_aggregation_map: Dict[Tuple[Any, Any, None, Optional[str]], Dict[str, decimal.Decimal]]
) -> Dict[Tuple[Any, Any, None, Optional[str]], Dict[str, decimal.Decimal]]:
    """
    CUSTOM Aggregation: Aggregates 'sqft' and 'amount' values based on unique
    combinations of 'po', 'item', AND 'description'. Uses a 4-element key
    (PO, Item, None, Description) for structural consistency with standard aggregation.
    Updates the global_custom_aggregation_map in place.
    """
    prefix = "[aggregate_custom]"
    required_cols = ['col_po', 'col_item', 'col_qty_sf', 'col_amount']

    if not processed_data:
        logging.info(f"{prefix} No data rows found in this table. Global custom aggregation map remains unchanged.")
        return global_custom_aggregation_map

    # Check for required columns existing in at least one row
    missing_cols = [col for col in required_cols if not any(col in row for row in processed_data)]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform full CUSTOM aggregation: Missing required columns {missing_cols}. Proceeding cautiously.")

    # Key selector
    def custom_key_selector(row):
        po_val, item_val = row.get('col_po'), row.get('col_item')
        desc_raw = row.get('col_desc')
        
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw
        description_key = description_key if description_key else None

        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"

        return (po_key, item_key, None, description_key)

    # Reducers
    reducers = {
        'col_qty_sf': ('col_qty_sf', decimal_sum_reducer),
        'col_amount': ('col_amount', decimal_sum_reducer),
        'col_net': ('col_net', decimal_sum_reducer),
        'col_cbm': ('col_cbm', decimal_sum_reducer)
    }

    # Run Aggregator
    aggregator = Aggregator(key_selector=custom_key_selector, reducers=reducers)
    aggregator.aggregate(processed_data, global_custom_aggregation_map)

    logging.info(f"{prefix} Finished processing rows.")
    return global_custom_aggregation_map


def calculate_leather_summary(processed_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Calculates the leather summary (PCS, SQFT, Net, Gross, Pallet Count) per leather type.
    Iterates through rows to sum values for each leather type found in 'description' or 'desc'.
    BUFFALO = rows containing "BUFFALO" in description
    COW = rows NOT containing "BUFFALO" (all other leather)
    """
    summary = {
        'BUFFALO': {'col_qty_pcs': 0, 'col_qty_sf': decimal.Decimal(0), 'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0), 'col_cbm': decimal.Decimal(0), 'col_pallet_count': 0},
        'COW': {'col_qty_pcs': 0, 'col_qty_sf': decimal.Decimal(0), 'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0), 'col_cbm': decimal.Decimal(0), 'col_pallet_count': 0}
    }

    if not processed_data:
        return summary

    def leather_type_key_selector(row):
        desc = str(row.get('col_desc', "")).upper() if row.get('col_desc') else ""
        return 'BUFFALO' if "BUFFALO" in desc else 'COW'

    reducers = {
        'col_qty_pcs': ('col_qty_pcs', int_sum_reducer),
        'col_qty_sf': ('col_qty_sf', decimal_sum_reducer),
        'col_net': ('col_net', decimal_sum_reducer),
        'col_gross': ('col_gross', decimal_sum_reducer),
        'col_cbm': ('col_cbm', decimal_sum_reducer),
        'col_pallet_count': ('col_pallet_count', int_sum_reducer),
    }

    aggregator = Aggregator(key_selector=leather_type_key_selector, reducers=reducers)
    aggregator.aggregate(processed_data, summary)
    return summary


def aggregate_per_po_with_pallets(processed_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Aggregates data by PO and Item, summing pallet, pcs, sqft, amount, net, gross, cbm.
    Groups rows that share the same (PO, Item) combination.
    """
    if not isinstance(processed_data, list) or not processed_data:
        return []

    # Key selector
    def po_item_key_selector(row):
        po_val = row.get('col_po')
        if po_val is None:
            return None
        po = str(po_val).strip()
        if not po:
            return None
        item_val = row.get('col_item')
        item = str(item_val).strip() if item_val is not None else ""
        return (po, item)

    reducers = {
        'col_desc': ('col_desc', first_non_empty_reducer),
        'col_qty_pcs': ('col_qty_pcs', int_sum_reducer),
        'col_qty_sf': ('col_qty_sf', decimal_sum_reducer),
        'col_amount': ('col_amount', decimal_sum_reducer),
        'col_pallet_count': ('col_pallet_count', int_sum_reducer),
        'col_net': ('col_net', decimal_sum_reducer),
        'col_gross': ('col_gross', decimal_sum_reducer),
        'col_cbm': ('col_cbm', decimal_sum_reducer),
    }

    # Aggregate using Aggregator
    aggregator = Aggregator(key_selector=po_item_key_selector, reducers=reducers)
    agg_map = aggregator.aggregate(processed_data)

    # Convert to list of dicts
    result = []
    for (po, item), data in agg_map.items():
        result.append({
            'col_po': po,
            'col_item': item,
            'col_desc': data.get('col_desc', ''),
            'col_qty_pcs': data.get('col_qty_pcs', 0),
            'col_qty_sf': data.get('col_qty_sf'),
            'col_amount': data.get('col_amount'),
            'col_pallet_count': data.get('col_pallet_count', 0),
            'col_net': data.get('col_net'),
            'col_gross': data.get('col_gross'),
            'col_cbm': data.get('col_cbm'),
        })

    # Sort by PO, then by Item for consistent output
    result.sort(key=lambda x: (x['col_po'], x['col_item']))
    
    logging.info(f"[aggregate_per_po_with_pallets] Aggregated {len(processed_data)} rows into {len(result)} unique PO+Item combinations.")
    return result
