import decimal
from typing import List, Any


def int_sum_reducer(values: List[Any]) -> int:
    """Sums integer values, assuming they are pre-normalized."""
    return sum(val for val in values if val is not None)


def decimal_sum_reducer(values: List[Any]) -> decimal.Decimal:
    """Sums numeric values as Decimal. Converts int/float/str to Decimal, skips unconvertible."""
    total = decimal.Decimal(0)
    for val in values:
        if val is None:
            continue
        if isinstance(val, decimal.Decimal):
            total += val
        else:
            try:
                total += decimal.Decimal(str(val))
            except (decimal.InvalidOperation, ValueError):
                pass
    return total


def first_non_empty_reducer(values: List[Any]) -> str:
    """Returns the first non-empty string value."""
    for val in values:
        if val:
            val_str = str(val).strip()
            if val_str:
                return val_str
    return ""
