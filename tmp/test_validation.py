
import decimal
import sys
import os

# Add root to path so we can import core
sys.path.append(os.getcwd())

from core.data_parser.validation import validate_weight_integrity, DataValidationError

# Test cases for Weight Integrity
print("Testing Weight Integrity...")

# 1. SUCCESS: Both present and Gross > Net
data_ok = [{'col_po': 'P1', 'col_item': 'I1', 'col_net': 100, 'col_gross': 110}]
try:
    validate_weight_integrity(data_ok)
    print("  [PASS] Standard pair valid")
except Exception as e:
    print(f"  [FAIL] Should have passed: {e}")

# 2. FAILURE: Missing Gross
data_no_gross = [{'col_po': 'P1', 'col_item': 'I1', 'col_net': 100}]
try:
    validate_weight_integrity(data_no_gross)
    print("  [FAIL] Should have raised error for missing Gross")
except DataValidationError as e:
    print(f"  [PASS] Caught missing Gross: {e}")

# 3. FAILURE: Missing Net
data_no_net = [{'col_po': 'P1', 'col_item': 'I1', 'col_gross': 110}]
try:
    validate_weight_integrity(data_no_net)
    print("  [FAIL] Should have raised error for missing Net")
except DataValidationError as e:
    print(f"  [PASS] Caught missing Net: {e}")

# 4. SUCCESS: Both missing (Skip)
data_empty = [{'col_po': 'P1', 'col_item': 'I1'}]
try:
    validate_weight_integrity(data_empty)
    print("  [PASS] Both missing (Skipped)")
except Exception as e:
    print(f"  [FAIL] Should have skipped: {e}")

# 5. FAILURE: Gross <= Net
data_bad_values = [{'col_po': 'P1', 'col_item': 'I1', 'col_net': 100, 'col_gross': 90}]
try:
    validate_weight_integrity(data_bad_values)
    print("  [FAIL] Should have caught Gross <= Net")
except DataValidationError as e:
    print(f"  [PASS] Caught Gross <= Net: {e}")

print("\nDone.")
