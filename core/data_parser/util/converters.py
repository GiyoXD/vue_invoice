import decimal
import logging
from typing import Any, Optional


class DataConverter:
    """
    A utility class that groups related data conversion functions.
    Methods are static as they do not depend on the state of an instance.
    """
    @staticmethod
    def convert_pallet_string(pallet_string: str) -> int:
        """
        Converts a pallet string into a pallet count based on specific rules.

        The logic is as follows:
        - A range 'x-y' (e.g., '1-2', '3-5', '2-2') counts as 1 pallet.
        - A single number (e.g., '5') counts as 1 pallet.
        - '0' or an empty/invalid string counts as 0 pallets.

        Args:
            pallet_string (str): The string representing the pallet notation.

        Returns:
            int: The calculated number of pallets (0 or 1).
        """
        if not pallet_string or not isinstance(pallet_string, str):
            return 0

        pallet_string = pallet_string.strip()
        
        if not pallet_string:
            return 0

        if '-' in pallet_string:
            try:
                start, end = map(int, pallet_string.split('-'))
                return 1
            except (ValueError, IndexError):
                return 0
        else:
            try:
                num = int(pallet_string)
                return 1 if num > 0 else 0
            except ValueError:
                return 0

    @staticmethod
    def convert_to_decimal(value: Any, context: str = "") -> Optional[decimal.Decimal]:
        """
        Safely convert a value to Decimal, logging warnings for common conversion issues.
        
        Args:
            value: The input value (likely float, int, str, or Decimal).
            context: Additional info (like row/col) for logging only.

        Returns:
            Optional[Decimal]: The converted decimal or None if conversion failed/invalid.
        """
        prefix = "[DataConverter.convert_to_decimal]"
        if isinstance(value, decimal.Decimal):
            return value
        if value is None:
            return None
        
        # Handle floats specially to avoid floating-point precision issues
        # repr() in Python 3.1+ gives the SHORTEST string that round-trips back
        # to the same float, e.g. repr(5028.2) → '5028.2'
        if isinstance(value, float):
            value_str = repr(value)
            if not value_str or value_str in ('-', 'nan', 'inf', '-inf'):
                return None
            try:
                # Set precision to ensure reliable conversion from repr string
                return decimal.Decimal(value_str)
            except (decimal.InvalidOperation, TypeError, ValueError) as e:
                logging.warning(f"{prefix} Could not convert float '{value}' to Decimal {context}: {e}")
                return None
        
        value_str = str(value).strip().replace(',', '')
        if not value_str:
            return None
        try:
            return decimal.Decimal(value_str)
        except (decimal.InvalidOperation, TypeError, ValueError) as e:
            logging.warning(f"{prefix} Could not convert '{value}' (Str: '{value_str}') to Decimal {context}: {e}")
            return None
