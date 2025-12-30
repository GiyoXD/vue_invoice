"""
Centralized definition of reusable style objects for the invoice generator.
"""

from openpyxl.styles import Alignment, Border, Side, Font

# --- Border Styles ---
THIN_SIDE = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
SIDE_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE)
NO_BORDER = Border(left=None, right=None, top=None, bottom=None)

# --- Alignment Styles ---
CENTER_ALIGNMENT = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_ALIGNMENT = Alignment(horizontal='left', vertical='center', wrap_text=True)

# --- Font Styles ---
BOLD_FONT = Font(bold=True)

# --- Constants for Number Formats ---
FORMAT_GENERAL = 'General'
FORMAT_TEXT = '@'
FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0'
FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00'
