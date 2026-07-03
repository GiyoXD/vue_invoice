import decimal
from typing import List, Any


def int_sum_reducer(values: List[Any]) -> int:
    """Sums integer values, assuming they are pre-normalized."""
    return sum(val for val in values if val is not None)


def decimal_sum_reducer(values: List[Any]) -> decimal.Decimal:
    """Sums decimal.Decimal values. Silently skips non-Decimal entries (e.g. raw strings)."""
    return sum((val for val in values if isinstance(val, decimal.Decimal)), decimal.Decimal(0))


def first_non_empty_reducer(values: List[Any]) -> str:
    """Returns the first non-empty string value."""
    for val in values:
        if val:
            val_str = str(val).strip()
            if val_str:
                return val_str
    return ""
