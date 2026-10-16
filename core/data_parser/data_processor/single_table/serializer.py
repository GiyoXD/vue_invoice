from decimal import Decimal
from typing import List, Dict, Any

__all__ = ["format_aggregation_as_list"]


def format_aggregation_as_list(
    aggregation_map: Dict[Any, Dict[str, Any]],
    convert_decimals: bool = False,
) -> List[Dict[str, Any]]:
    """
    Converts the internal aggregation map into a clean list of dictionaries
    suitable for JSON output.

    :param aggregation_map: Mapping containing aggregated row dictionaries.
    :param convert_decimals: If True, convert decimal.Decimal values to float
                             for standard json.dumps compatibility. If False
                             (default), preserve Decimal instances.
    :return: List of record dictionaries.
    """
    if convert_decimals:
        return [
            {k: float(v) if isinstance(v, Decimal) else v for k, v in values.items()}
            for values in aggregation_map.values()
        ]
    return [dict(values) for values in aggregation_map.values()]

