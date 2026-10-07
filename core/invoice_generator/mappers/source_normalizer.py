from typing import Any, Dict, List, Optional


def normalize_source_to_rows(data_source: Any) -> List[Dict[str, Any]]:
    """
    Convert raw input into a uniform List[Dict[str, Any]].
    Supports list of row dicts or a single row dict.
    """
    if not data_source:
        return []

    if isinstance(data_source, list):
        return [dict(item) for item in data_source if isinstance(item, dict)]

    if isinstance(data_source, dict):
        return [dict(data_source)]

    return []


def extract_cell_value(
    row: Dict[str, Any],
    rule_key: str,
    rule: Optional[Dict[str, Any]] = None,
) -> Any:
    """Extract value from a single normalized row dict."""
    if not isinstance(row, dict) or not rule_key:
        return None

    mapped = rule.get("column") if isinstance(rule, dict) else None
    primary = mapped or rule_key
    val = row.get(primary)
    if val not in (None, ""):
        return val

    if primary != rule_key:
        val = row.get(rule_key)
        if val not in (None, ""):
            return val

    return None


__all__ = ["normalize_source_to_rows", "extract_cell_value"]
