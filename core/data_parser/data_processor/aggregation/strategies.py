import logging
import decimal
from typing import List, Dict, Any, Tuple, Optional, Union
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
    Aggregates data by PO, Item, Unit Price, and Description (same as STANDARD aggregation),
    summing pallet, pcs, sqft, amount, net, gross, cbm.
    """
    if not isinstance(processed_data, list) or not processed_data:
        return []

    # Key selector (same as standard aggregation)
    def po_item_price_desc_key_selector(row):
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

    reducers = {
        'col_qty_pcs': ('col_qty_pcs', int_sum_reducer),
        'col_qty_sf': ('col_qty_sf', decimal_sum_reducer),
        'col_amount': ('col_amount', decimal_sum_reducer),
        'col_pallet_count': ('col_pallet_count', int_sum_reducer),
        'col_net': ('col_net', decimal_sum_reducer),
        'col_gross': ('col_gross', decimal_sum_reducer),
        'col_cbm': ('col_cbm', decimal_sum_reducer),
    }

    # Aggregate using Aggregator
    aggregator = Aggregator(key_selector=po_item_price_desc_key_selector, reducers=reducers)
    agg_map = aggregator.aggregate(processed_data)

    # Convert to list of dicts
    result = []
    for (po, item, price, desc), data in agg_map.items():
        result.append({
            'col_po': po,
            'col_item': item,
            'col_unit_price': str(price) if price is not None else "",
            'col_desc': desc if desc else "",
            'col_qty_pcs': data.get('col_qty_pcs', 0),
            'col_qty_sf': data.get('col_qty_sf'),
            'col_amount': data.get('col_amount'),
            'col_pallet_count': data.get('col_pallet_count', 0),
            'col_net': data.get('col_net'),
            'col_gross': data.get('col_gross'),
            'col_cbm': data.get('col_cbm'),
        })

    # Sort by PO, Item, and unit price for consistent output (natural/numeric order)
    def natural_sort_key(val):
        if val is None or val == "":
            return (0, "")
        try:
            return (1, float(val))
        except (ValueError, TypeError):
            return (2, str(val))

    result.sort(key=lambda x: (
        natural_sort_key(x['col_po']),
        natural_sort_key(x['col_item']),
        natural_sort_key(x['col_unit_price'])
    ))
    
    logging.info(f"[aggregate_per_po_with_pallets] Aggregated {len(processed_data)} rows into {len(result)} unique PO+Item+Price combinations.")
    return result
DAFCompoundingResult = Dict[str, Union[str, decimal.Decimal]]
FinalDAFResultType = List[DAFCompoundingResult]


def perform_DAF_compounding(
    data: Union[List[Dict[str, Any]], List[List[Dict[str, Any]]]],
    daf_chunk_size: int = 2,
    daf_intra_separator: str = "/",
    daf_inter_separator: str = "\n"
) -> Optional[FinalDAFResultType]:
    """
    Performs DAF Compounding directly from processed table rows or a list of tables.
    - If description data IS present: Performs BUFFALO split (Groups "1" & "2").
      Uses chunk_size=2 and separator='/'.
    - If description data IS NOT present: Performs PO Count split (Groups "1", "2", ...).
      Calculates chunk-specific totals.
    """
    prefix = "[perform_DAF_compounding]"
    logging.info(f"{prefix} Starting DAF Compounding directly from row data.")

    # Flatten input to a list of row dicts
    rows = []
    if data:
        if isinstance(data[0], list):
            for table in data:
                if isinstance(table, list):
                    rows.extend(table)
        else:
            rows = data

    # Helper function for creating a default empty group result
    def default_group_result() -> DAFCompoundingResult:
        return {
            'col_po': '',
            'col_item': '',
            'col_desc': '',
            'col_qty_sf': decimal.Decimal(0),
            'col_amount': decimal.Decimal(0)
        }

    # Handle empty input consistently -> returns default BUFFALO split dict
    if not rows:
        logging.warning(f"{prefix} Input rows list is empty. Returning default empty DAF groups.")
        return [
            default_group_result(), # Buffalo group
            default_group_result()  # Non-Buffalo group
        ]

    # --- Check if any description data exists ---
    any_description_present = any(
        row.get('col_desc') and str(row.get('col_desc')).strip()
        for row in rows
    )

    # Reusable helper function for formatting chunks
    def format_chunks(items: List[str], chunk_size: int, intra_sep: str, inter_sep: str) -> str:
        if not items:
            return ""
        processed_chunks = []
        for i in range(0, len(items), chunk_size):
            chunk = [str(item) for item in items[i:i + chunk_size]]
            joined_chunk = intra_sep.join(chunk)
            processed_chunks.append(joined_chunk)
        return inter_sep.join(processed_chunks)

    # --- Decide Execution Path --- #

    if any_description_present:
        # --- Path 1: Descriptions ARE present -> BUFFALO Split Aggregation ---
        logging.info(f"{prefix} Performing BUFFALO split aggregation (Chunk Size: {daf_chunk_size}).")
        # Initialize accumulators for BUFFALO group ("1")
        buffalo_pos = set()
        buffalo_items = set()
        buffalo_descriptions = set()
        buffalo_sqft = decimal.Decimal(0)
        buffalo_amount = decimal.Decimal(0)
        buffalo_net = decimal.Decimal(0)
        # Initialize accumulators for NON-BUFFALO group ("2")
        non_buffalo_pos = set()
        non_buffalo_items = set()
        non_buffalo_descriptions = set()
        non_buffalo_sqft = decimal.Decimal(0)
        non_buffalo_amount = decimal.Decimal(0)
        non_buffalo_net = decimal.Decimal(0)

        for row in rows:
            po_val = row.get('col_po')
            item_val = row.get('col_item')
            desc_val = row.get('col_desc')

            po_str = str(po_val).strip() if po_val is not None else "<MISSING_PO>"
            item_str = str(item_val).strip() if item_val is not None else "<MISSING_ITEM>"
            desc_str = str(desc_val).strip() if desc_val is not None else ""

            is_buffalo = desc_str and "BUFFALO" in desc_str.upper()

            # Use new col_ keys for sums
            sqft_sum = row.get('col_qty_sf', decimal.Decimal(0))
            amount_sum = row.get('col_amount', decimal.Decimal(0))
            net_sum = row.get('col_net', decimal.Decimal(0))

            if not isinstance(sqft_sum, decimal.Decimal): sqft_sum = decimal.Decimal(0)
            if not isinstance(amount_sum, decimal.Decimal): amount_sum = decimal.Decimal(0)
            if not isinstance(net_sum, decimal.Decimal): net_sum = decimal.Decimal(0)

            if is_buffalo:
                buffalo_pos.add(po_str)
                buffalo_items.add(item_str)
                buffalo_descriptions.add(desc_str)
                buffalo_sqft += sqft_sum
                buffalo_amount += amount_sum
                buffalo_net += net_sum
            else:
                non_buffalo_pos.add(po_str)
                non_buffalo_items.add(item_str)
                if desc_str: non_buffalo_descriptions.add(desc_str)
                non_buffalo_sqft += sqft_sum
                non_buffalo_amount += amount_sum
                non_buffalo_net += net_sum

        logging.debug(f"{prefix} Finished processing entries for BUFFALO split.")

        # Format BUFFALO Group ("1")
        sorted_buffalo_pos = sorted(list(buffalo_pos))
        sorted_buffalo_items = sorted(list(buffalo_items))
        sorted_buffalo_descriptions = sorted([d for d in buffalo_descriptions if d])
        buffalo_result: DAFCompoundingResult = {
            'col_po': format_chunks(sorted_buffalo_pos, daf_chunk_size, daf_intra_separator, daf_inter_separator),
            'col_item': format_chunks(sorted_buffalo_items, daf_chunk_size, daf_intra_separator, daf_inter_separator),
            'col_desc': format_chunks(sorted_buffalo_descriptions, 1, "", "\n"),
            'col_qty_sf': buffalo_sqft,
            'col_amount': buffalo_amount,
            'col_net': buffalo_net
        }
        # Format NON-BUFFALO Group ("2")
        sorted_non_buffalo_pos = sorted(list(non_buffalo_pos))
        sorted_non_buffalo_items = sorted(list(non_buffalo_items))
        sorted_non_buffalo_descriptions = sorted([d for d in non_buffalo_descriptions if d])
        non_buffalo_result: DAFCompoundingResult = {
            'col_po': format_chunks(sorted_non_buffalo_pos, daf_chunk_size, daf_intra_separator, daf_inter_separator),
            'col_item': format_chunks(sorted_non_buffalo_items, daf_chunk_size, daf_intra_separator, daf_inter_separator),
            'col_desc': format_chunks(sorted_non_buffalo_descriptions, 1, "", "\n"),
            'col_qty_sf': non_buffalo_sqft,
            'col_amount': non_buffalo_amount,
            'col_net': non_buffalo_net
        }
        # Construct Final Result LIST for BUFFALO Split Case
        final_buffalo_split_result: FinalDAFResultType = [
            buffalo_result,
            non_buffalo_result
        ]
        logging.info(f"{prefix} BUFFALO split DAF Compounding complete.")
        return final_buffalo_split_result
        # --- End Path 1 (BUFFALO Split) --- #

    else:
        # --- Path 2: Descriptions are NOT present -> Dynamic PO Count Split ---
        logging.info(f"{prefix} No description data found. Performing dynamic PO count split aggregation (split in half if >7 POs).")
        logging.info(f"{prefix}   - String formatting uses chunk size {daf_chunk_size} and separator '{daf_intra_separator}'.")

        # Step 1: Aggregate data by PO
        po_data_aggregation: Dict[str, Dict[str, Union[set, decimal.Decimal]]] = {}
        logging.debug(f"{prefix} Pass 1: Aggregating SQFT/Amount/Items per PO.")
        for row in rows:
            po_val = row.get('col_po')
            item_val = row.get('col_item')

            po_str = str(po_val).strip() if po_val is not None else "<MISSING_PO>"
            item_str = str(item_val).strip() if item_val is not None else "<MISSING_ITEM>"

            # Use new col_ keys
            sqft_sum = row.get('col_qty_sf', decimal.Decimal(0))
            amount_sum = row.get('col_amount', decimal.Decimal(0))
            net_sum = row.get('col_net', decimal.Decimal(0))

            if not isinstance(sqft_sum, decimal.Decimal): sqft_sum = decimal.Decimal(0)
            if not isinstance(amount_sum, decimal.Decimal): amount_sum = decimal.Decimal(0)
            if not isinstance(net_sum, decimal.Decimal): net_sum = decimal.Decimal(0)

            if po_str not in po_data_aggregation:
                po_data_aggregation[po_str] = {
                    'sqft_total': decimal.Decimal(0),
                    'amount_total': decimal.Decimal(0),
                    'net_total': decimal.Decimal(0),
                    'items': set()
                }
            po_data_aggregation[po_str]['sqft_total'] += sqft_sum # type: ignore
            po_data_aggregation[po_str]['amount_total'] += amount_sum # type: ignore
            po_data_aggregation[po_str]['net_total'] += net_sum # type: ignore
            po_data_aggregation[po_str]['items'].add(item_str) # type: ignore

        if not po_data_aggregation:
            logging.warning(f"{prefix} No valid PO data found for PO count splitting. Returning empty dict.")
            return []

        # Step 2: Get sorted list of unique POs
        sorted_pos = sorted(list(po_data_aggregation.keys()))

        # Step 3: Iterate through POs in chunks based on >7 rule for total calculation
        final_po_count_split_result: FinalDAFResultType = []
        
        # Determine conceptual chunks
        conceptual_po_chunks = []
        if len(sorted_pos) > 7:
            # Break down into half
            import math
            mid = math.ceil(len(sorted_pos) / 2)
            conceptual_po_chunks.append(sorted_pos[:mid])
            conceptual_po_chunks.append(sorted_pos[mid:])
            logging.debug(f"{prefix} Pass 2: Over 7 POs detected ({len(sorted_pos)}). Splitting into 2 chunks of {mid} and {len(sorted_pos)-mid}.")
        else:
            conceptual_po_chunks.append(sorted_pos)
            logging.debug(f"{prefix} Pass 2: {len(sorted_pos)} POs detected (<= 7). Keeping as 1 chunk.")

        for i, conceptual_po_chunk in enumerate(conceptual_po_chunks):
            # Calculate totals and collect items for THIS conceptual chunk
            chunk_sqft_total = decimal.Decimal(0)
            chunk_amount_total = decimal.Decimal(0)
            chunk_net_total = decimal.Decimal(0)
            chunk_items = set()
            po_list_for_formatting = [] # Collect POs in this chunk for formatting

            for po_str in conceptual_po_chunk:
                po_agg_data = po_data_aggregation.get(po_str)
                if po_agg_data:
                    chunk_sqft_total += po_agg_data.get('sqft_total', decimal.Decimal(0)) # type: ignore
                    chunk_amount_total += po_agg_data.get('amount_total', decimal.Decimal(0)) # type: ignore
                    chunk_net_total += po_agg_data.get('net_total', decimal.Decimal(0)) # type: ignore
                    chunk_items.update(po_agg_data.get('items', set())) # type: ignore
                    po_list_for_formatting.append(po_str) # Add the PO itself to the list for formatting
                else:
                     logging.warning(f"{prefix} PO '{po_str}' not found in aggregation data during chunking.")

            # Sort items collected for this chunk
            sorted_chunk_items = sorted(list(chunk_items))

            # Step 4: Format the collected POs and Items using desired format (size 2)
            formatted_po_chunk = format_chunks(po_list_for_formatting, daf_chunk_size, daf_intra_separator, daf_inter_separator)
            formatted_item_chunk = format_chunks(sorted_chunk_items, daf_chunk_size, daf_intra_separator, daf_inter_separator)

            # Create the result dictionary for this chunk index
            chunk_result: DAFCompoundingResult = {
                'col_po': formatted_po_chunk,
                'col_item': formatted_item_chunk,
                'col_desc': '', # No descriptions in this path
                'col_qty_sf': chunk_sqft_total,    # Use CHUNK total (calculated based on group of 8)
                'col_amount': chunk_amount_total,   # Use CHUNK total (calculated based on group of 8)
                'col_net': chunk_net_total
            }
            chunk_index_str = str(i + 1)
            final_po_count_split_result.append(chunk_result)
            logging.debug(f"{prefix} Created output chunk {chunk_index_str}: {len(conceptual_po_chunk)} POs contributed totals, SQFT={chunk_sqft_total}, Amount={chunk_amount_total}, Net={chunk_net_total}")

        logging.info(f"{prefix} PO count split DAF Compounding complete ({len(final_po_count_split_result)} chunks created).")
        return final_po_count_split_result
        # --- End Path 2 (PO Count Split) --- #
