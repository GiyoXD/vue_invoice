import decimal
from typing import List, Any

__all__ = [
    "int_sum_reducer",
    "decimal_sum_reducer",
    "first_value_reducer",
    "first_non_empty_reducer",
    "unique_join_reducer",
    "po_merge_reducer",
]


def int_sum_reducer(values: List[Any]) -> int:
    """Sums integer values, assuming they are pre-normalized."""
    return sum(val for val in values if val is not None)


def first_value_reducer(values: List[Any]) -> Any:
    """Returns the first non-None value."""
    for val in values:
        if val is not None:
            return val
    return None


def decimal_sum_reducer(values: List[Any]) -> decimal.Decimal:
    """Sums numeric values as Decimal."""
    total = decimal.Decimal(0)
    for v in values:
        if v is None or v is True or v is False:
            continue
        try:
            total += v
        except TypeError:
            try:
                clean_str = str(v).replace(",", "").strip()
                total += decimal.Decimal(clean_str)
            except (decimal.InvalidOperation, ValueError, TypeError):
                pass
    return total


def first_non_empty_reducer(values: List[Any]) -> str:
    """Returns the first non-empty string value."""
    for val in values:
        if val is not None:
            val_str = str(val).strip()
            if val_str:
                return val_str
    return ""


def unique_join_reducer(values: List[Any]) -> str:
    """Merges distinct strings, splitting on ' / ' and rejoining in sorted order."""
    unique_pos = set()
    for val in values:
        if val is None:
            continue
        val_str = str(val).strip()
        if not val_str:
            continue
        for part in val_str.split(" / "):
            p = part.strip()
            if p:
                unique_pos.add(p)
    return " / ".join(sorted(unique_pos)) if unique_pos else ""


po_merge_reducer = unique_join_reducer
