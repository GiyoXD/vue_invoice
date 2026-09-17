from typing import List, Dict, Any

__all__ = ["format_aggregation_as_list"]


def format_aggregation_as_list(
    aggregation_map: Dict[Any, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Converts the internal aggregation map into a clean list of dictionaries
    suitable for JSON output.
    """
    return [dict(values) for values in aggregation_map.values()]
