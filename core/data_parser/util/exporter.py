import datetime
import json
import logging
import os
import shutil
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Optional

from .serializer import Serializer
from .. import data_processor
from .. import config as cfg

logger = logging.getLogger(__name__)


def export_invoice_data(
    output_dir: Path,
    input_filename: str,
    input_stem: str,
    actual_sheet_name: str,
    aggregation_mode_used: str,
    processed_tables: List[List[Dict[str, Any]]],
    table_footer_data: List[Dict[str, Any]],
    grand_total_footer: Dict[str, Any],
    leather_summary: Dict[str, Any],
    weight_summary_addon: Dict[str, Any],
    global_standard_aggregation_results: Dict,
    global_custom_aggregation_results: Dict,
    normal_aggregate_per_po: List[Dict[str, Any]],
    global_DAF_compounded_result: Any,
    warnings: List[str]
) -> Path:
    """
    Constructs the final JSON payload from the processed invoice data,
    serializes it using the custom Serializer utility, and atomically writes it to disk.

    Args:
        output_dir (Path): Directory where the output JSON file will be stored.
        input_filename (str): The name of the processed input Excel file.
        input_stem (str): The filename stem (no extension) of the input Excel file.
        actual_sheet_name (str): The title of the worksheet processed.
        aggregation_mode_used (str): The name of the aggregation mode used (standard/custom).
        processed_tables (List[List[Dict[str, Any]]]): Processed table rows containing the distributed/normalized values.
        table_footer_data (List[Dict[str, Any]]): Summed/footer values per table.
        grand_total_footer (Dict[str, Any]): Overall grand totals for all columns.
        leather_summary (Dict[str, Any]): Summary statistics (pallet count, weight) by leather type.
        weight_summary_addon (Dict[str, Any]): Merged weight totals across all tables.
        global_standard_aggregation_results (Dict): Aggregations using PO + Item + Price standard groupings.
        global_custom_aggregation_results (Dict): Aggregations using PO + Item custom groupings.
        normal_aggregate_per_po (List[Dict[str, Any]]): Aggregated manifest partitioned by PO & price with accurate pallet ranges.
        global_DAF_compounded_result (Any): Compounded DAF manifest groupings.
        warnings (List[str]): Run-time processing warnings.

    Returns:
        Path: The absolute path to the successfully written and verified JSON output file.
    """
    # 1. Construct final payload structure
    final_json_structure = {
        "metadata": {
            "workbook_filename": input_filename,
            "worksheet_name": actual_sheet_name,
            "DAF_compounding_input_mode": aggregation_mode_used,
            "DAF_chunk_size": cfg.DAF_CHUNK_SIZE,
            "DAF_intra_separator": cfg.DAF_INTRA_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'),
            "DAF_inter_separator": cfg.DAF_INTER_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'),
            "timestamp": datetime.datetime.now().isoformat(),
            "warnings": warnings
        },
        "price_adjustment": [],
        "multi_table": processed_tables,
        "footer_data": {
            "table_totals": table_footer_data,
            "grand_total": grand_total_footer,
            "add_ons": {
                "leather_summary_addon": leather_summary,
                "weight_summary_addon": weight_summary_addon,
            }
        },
        "single_table": {
            "aggregation": data_processor.format_aggregation_as_list(global_standard_aggregation_results, mode='standard'),
            "aggregation_custom": data_processor.format_aggregation_as_list(global_custom_aggregation_results, mode='custom'),
            "manifest_by_pallet_per_po": normal_aggregate_per_po,
            "aggregation_DAF": global_DAF_compounded_result
        }
    }

    # 2. Serialize payload using Serializer helper
    try:
        json_output_string = Serializer.serialize_to_json(final_json_structure, indent=4)
    except TypeError as json_err:
        logger.error(f"Failed to serialize data to JSON: {json_err}. Check data types.", exc_info=True)
        raise json_err
    except Exception as e:
        logger.error(f"An unexpected error occurred during JSON generation: {e}", exc_info=True)
        raise e

    logger.info(f"Generated JSON output structure successfully ({len(json_output_string)} chars).")

    # 3. Determine target filename & path
    json_output_filename = f"{input_stem}.json"
    output_json_path = output_dir / json_output_filename
    logger.info(f"Determined output JSON path: {output_json_path}")

    # 4. Atomic Write with post-write verification
    try:
        temp_fd, temp_path = tempfile.mkstemp(
            suffix='.json.tmp', dir=str(output_json_path.parent)
        )
        try:
            with os.fdopen(temp_fd, 'w', encoding='utf-8') as f_json:
                f_json.write(json_output_string)
                f_json.flush()
                os.fsync(f_json.fileno())  # Guarantee persistence to media

            # Verify integrity of written output
            with open(temp_path, 'r', encoding='utf-8') as f_verify:
                json.load(f_verify)  # Raises JSONDecodeError if truncated or corrupted

            # Atomic swap
            shutil.move(temp_path, str(output_json_path))
            logger.info(f"Successfully saved JSON output to '{output_json_path}' (verified)")
        except Exception:
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            raise
    except json.JSONDecodeError as verify_err:
        logger.error(f"CRITICAL: JSON integrity check failed after write — output would be truncated/corrupt: {verify_err}")
        raise RuntimeError(
            f"JSON output verification failed: the generated data could not be re-parsed. "
            f"Details: {verify_err}"
        )
    except IOError as io_err:
        logger.error(f"Failed to write JSON output to file '{output_json_path}': {io_err}")
        raise io_err
    except Exception as write_err:
        logger.error(f"An unexpected error occurred while writing JSON file: {write_err}", exc_info=True)
        raise write_err

    return output_json_path
