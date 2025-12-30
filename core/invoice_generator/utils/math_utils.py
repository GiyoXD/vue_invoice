"""
Math Utilities

Provides robust functions for safely converting values to numbers (float/int),
handling common issues like whitespace, string representations, and negative numbers.
"""

import logging
from typing import Any, Optional, Union

logger = logging.getLogger(__name__)

def safe_float_convert(value: Any, default: float = 0.0) -> float:
    """
    Safely converts a value to a float.
    
    Handles:
    - Integers and floats (returned as float)
    - Strings with whitespace
    - Strings with negative signs
    - Strings with decimal points
    
    Args:
        value: The value to convert.
        default: The default value to return if conversion fails.
        
    Returns:
        The converted float value, or the default if conversion fails.
    """
    if value is None:
        return default
        
    if isinstance(value, (int, float)):
        return float(value)
        
    if isinstance(value, str):
        try:
            # Strip whitespace
            cleaned = value.strip()
            if not cleaned:
                return default
                
            # Check if it looks like a number (handling negative sign and decimal point)
            # We use a simple check first to avoid try/except overhead for obvious non-numbers
            # But float() is the ultimate validator
            return float(cleaned)
        except (ValueError, TypeError):
            # Log at debug level to avoid spamming logs for expected non-numeric cells
            # logger.debug(f"Failed to convert '{value}' to float. Using default {default}.")
            pass
            
    return default

def safe_int_convert(value: Any, default: int = 0) -> int:
    """
    Safely converts a value to an integer.
    
    Handles:
    - Integers (returned as int)
    - Floats (truncated to int)
    - Strings with whitespace
    - Strings with negative signs
    
    Args:
        value: The value to convert.
        default: The default value to return if conversion fails.
        
    Returns:
        The converted integer value, or the default if conversion fails.
    """
    if value is None:
        return default
        
    if isinstance(value, int):
        return value
        
    if isinstance(value, float):
        return int(value)
        
    if isinstance(value, str):
        try:
            # Strip whitespace
            cleaned = value.strip()
            if not cleaned:
                return default
                
            # Use float() first to handle strings like "10.5" -> 10.5 -> 10
            return int(float(cleaned))
        except (ValueError, TypeError):
            pass
            
    return default
