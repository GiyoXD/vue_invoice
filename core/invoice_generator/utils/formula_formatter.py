import re
from typing import Optional

def extract_decimal_places(number_format_str: Optional[str]) -> Optional[int]:
    if number_format_str is None:
        return None
    cleaned_fmt = re.sub(r'"[^"]*"', '', number_format_str)
    if '.' in cleaned_fmt:
        after_dot = cleaned_fmt.split('.')[1]
        digits = 0
        for char in after_dot:
            if char in ('0', '#', '?'):
                digits += 1
            elif char in (';', ' ', '_', '%', ']'):
                break
        return digits
    if any(c in cleaned_fmt for c in ('0', '#')):
        return 0
    return None

def wrap_with_round(formula: str, decimals: Optional[int]) -> str:
    if decimals is None or not isinstance(decimals, int):
        return formula
    if not formula:
        return formula
    expr = formula.strip()
    has_equals = expr.startswith('=')
    expr_body = expr[1:].strip() if has_equals else expr
    if expr_body.upper().startswith('ROUND('):
        return expr if has_equals else f"={expr}"
    return f"=ROUND({expr_body}, {decimals})"
