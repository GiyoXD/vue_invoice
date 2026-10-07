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


EXT_SHEET_REF_OR_LITERAL_PATTERN = re.compile(
    r'\"(?:\"\"|[^\"])*\"|'
    r'(?:\'(?:[^\']|\'\')+\'|[A-Za-z0-9_\[\]\.\:]+)\s*!\s*(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)|'
    r'\'(?:[^\']|\'\')*\''
)


def shift_formula_coordinates(
    formula: str,
    row_offset: int,
    header_end_row: int = 0,
    original_footer_start_row: int = 0
) -> str:
    """
    Shifts row coordinates in an Excel formula string based on row_offset.
    
    Args:
        formula: Formula string (e.g. '=SUM(B10:B20)').
        row_offset: Number of rows to shift (footer_start_row - original_footer_start_row).
        header_end_row: End row of header; references <= header_end_row are not shifted.
        original_footer_start_row: Start row of footer in original template.
        
    Returns:
        Updated formula string with shifted coordinates.
    """
    if row_offset == 0 or not formula or not isinstance(formula, str) or not formula.startswith('='):
        return formula

    header_end_row = header_end_row or 0
    original_footer_start_row = original_footer_start_row or 0

    # Mask string literals (e.g. "...") and external sheet references (e.g. 'Sheet'!A1 or Sheet!A1:B10) before regex parsing
    string_literals = []

    def mask_literal(m: re.Match) -> str:
        idx = len(string_literals)
        string_literals.append(m.group(0))
        return f"___STR_LITERAL_{idx}___"

    masked_formula = EXT_SHEET_REF_OR_LITERAL_PATTERN.sub(mask_literal, formula)

    pattern = re.compile(r'(?<![A-Za-z0-9_!])(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?![A-Za-z0-9_!\']|\s*\()')

    def replace_coord(match: re.Match) -> str:
        col_dollar = match.group(1)
        col_letters = match.group(2)
        row_dollar = match.group(3)
        row_str = match.group(4)
        row_num = int(row_str)

        if original_footer_start_row > 0:
            table_data_end_row = original_footer_start_row - 1
            if row_num != table_data_end_row and row_num < original_footer_start_row:
                return match.group(0)
        elif header_end_row > 0 and row_num <= header_end_row:
            return match.group(0)

        new_row = row_num + row_offset
        if new_row <= 0:
            return "#REF!"
        return f"{col_dollar}{col_letters}{row_dollar}{new_row}"

    shifted_formula = pattern.sub(replace_coord, masked_formula)

    if string_literals:
        def unmask_literal(m: re.Match) -> str:
            idx = int(m.group(1))
            return string_literals[idx]

        return re.sub(r'___STR_LITERAL_(\d+)___', unmask_literal, shifted_formula)

    return shifted_formula


