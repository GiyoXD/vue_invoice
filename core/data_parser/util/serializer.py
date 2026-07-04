import datetime
import decimal
import json
from typing import Any


class Serializer:
    """
    A utility class that groups serialization and conversion functions.
    Handles converting custom types (datetime, date, Decimal, set) and recursively
    stringifying dictionary keys (e.g. tuple keys) so they are JSON-serializable.
    """
    @staticmethod
    def _serialize_unsupported_value(obj: Any) -> Any:
        """JSON serializer for objects not serializable by default json code"""
        if isinstance(obj, (datetime.datetime, datetime.date)):
            return obj.isoformat()  # Convert date/datetime to ISO string format
        elif isinstance(obj, decimal.Decimal):  # Convert Decimal to float
            return float(obj)
        elif isinstance(obj, set):  # Convert set to list
            return list(obj)
        raise TypeError(f"Object of type {obj.__class__.__name__} is not JSON serializable")

    @classmethod
    def _stringify_keys_recursively(cls, data: Any) -> Any:
        """Recursively converts tuple keys in dicts to strings and handles non-serializable types."""
        if isinstance(data, dict):
            # Convert all keys to string, including tuple keys
            return {str(k): cls._stringify_keys_recursively(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [cls._stringify_keys_recursively(item) for item in data]
        elif data is None:
            return None  # JSON null
        return data

    @classmethod
    def serialize_to_json(cls, data: Any, indent: int = 4) -> str:
        """
        Preprocesses key structures (e.g. converting tuple keys to strings) and
        serializes the entire structure to JSON with fallback handlers.
        """
        clean_data = cls._stringify_keys_recursively(data)
        return json.dumps(
            clean_data,
            indent=indent,
            default=cls._serialize_unsupported_value
        )
