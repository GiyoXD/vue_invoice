import ast
import re
import logging
from typing import Any, Dict, List, Tuple, Union, Optional

logger = logging.getLogger(__name__)


def extract_table_data(data_source: Any, data_source_type: str) -> Any:
    """
    Extract data for the specific table being processed.
    
    For multi-table data sources, this extracts the subset for table_key.
    For single-table sources, returns the full data source.
    """
    if data_source is None:
        return None
    
    # For processed_tables_multi, BuilderConfigResolver already extracted the table
    if data_source_type in ['processed_tables', 'processed_tables_multi']:
        return data_source
    
    # For other types like aggregation, return as-is
    # Check for stringified tuple keys (JSON artifact) and convert back to tuples
    if isinstance(data_source, dict):
        new_data = {}
        for k, v in data_source.items():
            if isinstance(k, str) and k.startswith('(') and k.endswith(')'):
                try:
                    # Clean up Decimal wrappers for literal_eval: "Decimal('1.2')" -> "1.2"
                    clean_k = re.sub(r"Decimal\((['\"])(.*?)\1\)", r"\2", k)
                    new_key = ast.literal_eval(clean_k)
                    new_data[new_key] = v
                except (ValueError, SyntaxError):
                    new_data[k] = v
            else:
                new_data[k] = v
        return new_data

    return data_source
