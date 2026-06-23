import unittest
from decimal import Decimal
from core.data_parser import data_processor
from core.data_parser import sheet_parser
from core.data_parser.validation import validate_data, DataValidationError

class TestDataParserRefactor(unittest.TestCase):

    def setUp(self):
        # Current data structure representation
        self.mock_legacy_extract = {
            1: {
                "col_po": ["A1"], "col_item": ["I1"], "col_qty_sf": [Decimal('100.5')],
                "col_cbm": ["1x2x3"], "col_unit_price": [Decimal('1.5')], "col_amount": [Decimal('150.75')]
            }
        }

        # Another mock for distribution
        self.mock_legacy_distribute = {
            1: {
                "col_po": ["A1", "A1", "A1"],
                "col_item": ["I1", "I2", "I3"],
                "col_qty_sf": [Decimal('50'), Decimal('30'), Decimal('20')],
                "col_amount": [Decimal('100'), None, None]  # 100 to distribute among 50, 30, 20
            }
        }

        # Proposed PROSPECTIVE data structure representation
        self.mock_new_extract = [
            [
                {
                    "col_po": "A1", "col_item": "I1", "col_qty_sf": Decimal('100.5'),
                    "col_cbm": "1x2x3", "col_unit_price": Decimal('1.5'), "col_amount": Decimal('150.75')
                }
            ]
        ]
        
        self.mock_new_distribute = [
            [
                {"col_po": "A1", "col_item": "I1", "col_qty_sf": Decimal('50'), "col_amount": Decimal('100')},
                {"col_po": "A1", "col_item": "I2", "col_qty_sf": Decimal('30'), "col_amount": None},
                {"col_po": "A1", "col_item": "I3", "col_qty_sf": Decimal('20'), "col_amount": None}
            ]
        ]


    # --- OLD LEGACY BEHAVIOUR VERIFICATION TESTS (Now tests modern behavior) --- #
    def test_legacy_cbm(self):
        # We now pass List[Dict] to process_cbm_column
        processed = data_processor.process_cbm_column(list(self.mock_new_extract[0]))
        self.assertEqual(processed[0]['col_cbm'], Decimal('6.0000'))

    def test_legacy_distribute_values(self):
        processed = data_processor.distribute_values(
            list(self.mock_new_distribute[0]), 
            columns_to_distribute=['col_amount'], 
            basis_column='col_qty_sf'
        )
        self.assertEqual(processed[0]['col_amount'], Decimal('50.00'))
        self.assertEqual(processed[1]['col_amount'], Decimal('30.00'))
        self.assertEqual(processed[2]['col_amount'], Decimal('20.00'))

    def test_legacy_standard_aggregation(self):
        global_map = {}
        data = list(self.mock_new_extract[0])
        # Add a duplicate row
        data.append(data[0].copy())
        
        res = data_processor.aggregate_standard_by_po_item_price(data, global_map)
        expected_key = ("A1", "I1", Decimal('1.5'), None)
        self.assertIn(expected_key, res)
        self.assertEqual(res[expected_key]['sqft_sum'], Decimal('201.0'))
        self.assertEqual(res[expected_key]['amount_sum'], Decimal('301.50'))

    # --- PROSPECTIVE BEHAVIOUR TESTS --- #
    def test_new_cbm(self):
        # Once process_cbm_column is refactored, it should accept a List[Dict] (the table rows)
        # Note: we are designing the test for the future function definition
        try:
             processed = data_processor.process_cbm_column(list(self.mock_new_extract[0]))
             self.assertEqual(processed[0]['col_cbm'], Decimal('6.0000'))
        except TypeError as err:
             self.fail(f"New schema failed: {err}. Function Needs Refactoring!")
        except Exception as err:
             self.fail(f"New schema logic failed: {err}")

    def test_new_distribute_values(self):
        try:
            processed = data_processor.distribute_values(
                list(self.mock_new_distribute[0]),
                columns_to_distribute=['col_amount'],
                basis_column='col_qty_sf'
            )
            self.assertEqual(processed[0]['col_amount'], Decimal('50.0000'))
            self.assertEqual(processed[1]['col_amount'], Decimal('30.0000'))
            self.assertEqual(processed[2]['col_amount'], Decimal('20.0000'))
        except TypeError as err:
             self.fail(f"New schema failed: {err}. Function Needs Refactoring!")
        except Exception as err:
             self.fail(f"New schema logic failed: {err}")
             

    def test_new_standard_aggregation(self):
        global_map = {}
        data = list(self.mock_new_extract[0])
        # Add duplicate row
        data.append(data[0].copy())
        try:
            res = data_processor.aggregate_standard_by_po_item_price(data, global_map)
            expected_key = ("A1", "I1", Decimal('1.5'), None)
            self.assertIn(expected_key, res)
            self.assertEqual(res[expected_key]['sqft_sum'], Decimal('201.0'))
            self.assertEqual(res[expected_key]['amount_sum'], Decimal('301.50'))
        except TypeError as err:
             self.fail(f"New schema failed: {err}. Function Needs Refactoring!")
        except Exception as err:
             self.fail(f"New schema logic failed: {err}")

    def test_cbm_pcs_proportion_orphaned_row(self):
        # Row 1 has pieces (col_qty_pcs=100) but col_cbm=0 (orphaned)
        # Row 2 has col_qty_pcs=0 but col_cbm=1.5
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('0'), "_row_num": 12},
            {"col_po": "PO1", "col_item": "ITEM2", "col_qty_pcs": Decimal('0'), "col_cbm": Decimal('1.5'), "_row_num": 13}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}
        
        # Should raise DataValidationError
        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion')
        self.assertIn("[Zero CBM] (Row 12)", str(context.exception))
        self.assertIn("has pieces (100) but received 0 CBM", str(context.exception))

        # With ignore_cbm_warning=True, it should pass
        try:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion', ignore_cbm_warning=True)
        except DataValidationError:
            self.fail("validate_data raised DataValidationError when ignore_cbm_warning was True")

    def test_cbm_pcs_proportion_monotonicity_violation(self):
        # ITEM1 has col_qty_pcs=100 and col_cbm=1.0
        # ITEM2 has col_qty_pcs=50 and col_cbm=2.0
        # Monotonicity violation
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('1.0'), "_row_num": 14},
            {"col_po": "PO1", "col_item": "ITEM2", "col_qty_pcs": Decimal('50'), "col_cbm": Decimal('2.0'), "_row_num": 15}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}

        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion')
        self.assertIn("[Monotonicity] (Rows 14 and 15)", str(context.exception))
        self.assertIn("more pieces (100) but lower CBM (1.0)", str(context.exception))

        # With ignore_cbm_warning=True, it should pass
        try:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion', ignore_cbm_warning=True)
        except DataValidationError:
            self.fail("validate_data raised DataValidationError when ignore_cbm_warning was True")

    def test_cbm_pcs_proportion_monotonicity_tolerance(self):
        # ITEM1 has 100 pcs and 1.95 CBM
        # ITEM2 has 50 pcs and 2.0 CBM
        # Difference is 2.5%, which is within the 3.2% tolerance -> should pass
        data_valid = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('1.95'), "_row_num": 14},
            {"col_po": "PO1", "col_item": "ITEM2", "col_qty_pcs": Decimal('50'), "col_cbm": Decimal('2.0'), "_row_num": 15}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}
        try:
            validate_data(data_valid, "Table 1", column_mapping, phase='cbm_proportion')
        except DataValidationError as ve:
            self.fail(f"Monotonicity with 2.5% difference failed validation: {ve}")

        # ITEM1 has 100 pcs and 1.93 CBM
        # ITEM2 has 50 pcs and 2.0 CBM
        # Difference is 3.5%, which is outside the 3.2% tolerance -> should fail
        data_invalid = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('1.93'), "_row_num": 14},
            {"col_po": "PO1", "col_item": "ITEM2", "col_qty_pcs": Decimal('50'), "col_cbm": Decimal('2.0'), "_row_num": 15}
        ]
        with self.assertRaises(DataValidationError) as context:
            validate_data(data_invalid, "Table 1", column_mapping, phase='cbm_proportion')
        self.assertIn("[Monotonicity]", str(context.exception))

    def test_cbm_pcs_proportion_abnormally_high_ratio(self):
        # ITEM1 has col_qty_pcs=2 and col_cbm=1.5 -> ratio = 0.75 > 0.5
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('2'), "col_cbm": Decimal('1.5'), "_row_num": 16}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}

        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion')
        self.assertIn("[High Ratio] (Row 16)", str(context.exception))
        self.assertIn("Exceeds 0.5 CBM/unit threshold", str(context.exception))

        # With ignore_cbm_warning=True, it should pass
        try:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion', ignore_cbm_warning=True)
        except DataValidationError:
            self.fail("validate_data raised DataValidationError when ignore_cbm_warning was True")

    def test_cbm_pcs_proportion_ratio_outliers(self):
        # Row 5 ratio (0.005) is outlier compared to median of 0.1
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('10.0'), "_row_num": 17},
            {"col_po": "PO1", "col_item": "ITEM2", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('10.0'), "_row_num": 18},
            {"col_po": "PO1", "col_item": "ITEM3", "col_qty_pcs": Decimal('100'), "col_cbm": Decimal('10.0'), "_row_num": 19},
            {"col_po": "PO1", "col_item": "ITEM4", "col_qty_pcs": Decimal('10'), "col_cbm": Decimal('2.0'), "_row_num": 20},
            {"col_po": "PO1", "col_item": "ITEM5", "col_qty_pcs": Decimal('10'), "col_cbm": Decimal('0.05'), "_row_num": 21}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}

        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion')
        self.assertIn("[Ratio Outlier] (Row 21)", str(context.exception))
        self.assertIn("Verify CBM (0.05) or quantity (10)", str(context.exception))

        # With ignore_cbm_warning=True, it should pass
        try:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion', ignore_cbm_warning=True)
        except DataValidationError:
            self.fail("validate_data raised DataValidationError when ignore_cbm_warning was True")

    def test_cbm_pcs_proportion_valid_table(self):
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('10'), "col_cbm": Decimal('1.0')},
            {"col_po": "PO1", "col_item": "ITEM2", "col_qty_pcs": Decimal('20'), "col_cbm": Decimal('2.0')},
            {"col_po": "PO1", "col_item": "ITEM3", "col_qty_pcs": Decimal('30'), "col_cbm": Decimal('3.0')}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}

        try:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion')
        except DataValidationError as ve:
            self.fail(f"Valid table failed validation: {ve}")

    def test_weight_integrity_unpaired_empty_string(self):
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_net": Decimal('10.0'), "col_gross": "", "_row_num": 5}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_net": "C", "col_gross": "D"}
        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='integrity')
        self.assertIn("Weight Integrity Error (Row 5)", str(context.exception))
        self.assertIn("Partial weight found", str(context.exception))

    def test_weight_integrity_both_empty(self):
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_net": "", "col_gross": None, "_row_num": 5}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_net": "C", "col_gross": "D"}
        try:
            validate_data(data, "Table 1", column_mapping, phase='integrity')
        except DataValidationError as ve:
            self.fail(f"Empty weights failed validation: {ve}")

    def test_weight_integrity_invalid_positivity(self):
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_net": Decimal('10.0'), "col_gross": Decimal('9.0'), "_row_num": 6}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_net": "C", "col_gross": "D"}
        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='integrity')
        self.assertIn("Weight Validation Error (Row 6)", str(context.exception))
        self.assertIn("Gross Weight (9.0) is not strictly greater than Net Weight (10.0)", str(context.exception))

    def test_validate_data_runs_pallet_integrity(self):
        # Weight valid, but pallet ID reappears after gap
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_net": Decimal('10.0'), "col_gross": Decimal('11.0'), "col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 10},
            {"col_po": "PO1", "col_item": "ITEM2", "col_net": Decimal('20.0'), "col_gross": Decimal('21.0'), "col_pallet_count": 1, "col_pallet_id": "01T26052608", "_row_num": 11},
            {"col_po": "PO1", "col_item": "ITEM3", "col_net": Decimal('30.0'), "col_gross": Decimal('31.0'), "col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 12},
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_net": "C", "col_gross": "D", "col_pallet_count": "E", "col_pallet_id": "F"}
        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='integrity')
        self.assertIn("Pallet ID '01T26052605' reappeared after a gap", str(context.exception))

    def test_verify_pallet_integrity_valid(self):
        data = [
            {"col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 10},
            {"col_pallet_count": 1, "col_pallet_id": "01T26052608", "_row_num": 11},
            {"col_pallet_count": 1, "col_pallet_id": "01T26052609", "_row_num": 12},
            {"col_pallet_count": 0, "col_pallet_id": "01T26052609", "_row_num": 13},
            {"col_pallet_count": 1, "col_pallet_id": "02T26052306", "_row_num": 14},
            {"col_pallet_count": 1, "col_pallet_id": "02T26052307", "_row_num": 15},
            {"col_pallet_count": 0, "col_pallet_id": "02T26052307", "_row_num": 16},
        ]
        # Should not raise any error
        try:
            data_processor.verify_pallet_integrity(data)
        except Exception as e:
            self.fail(f"verify_pallet_integrity raised an error on valid data: {e}")

    def test_verify_pallet_integrity_invalid_transition(self):
        # Pallet ID changed from 01T26052605 to 01T26052608, but count is 0
        data = [
            {"col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 10},
            {"col_pallet_count": 0, "col_pallet_id": "01T26052608", "_row_num": 11},
        ]
        with self.assertRaises(DataValidationError) as context:
            data_processor.verify_pallet_integrity(data)
        self.assertIn("Pallet ID changed to '01T26052608'", str(context.exception))
        self.assertIn("boundary marker count is 0", str(context.exception))

    def test_verify_pallet_integrity_invalid_continuity(self):
        # Pallet ID remains 01T26052609, but count is 1
        data = [
            {"col_pallet_count": 1, "col_pallet_id": "01T26052609", "_row_num": 12},
            {"col_pallet_count": 1, "col_pallet_id": "01T26052609", "_row_num": 13},
        ]
        with self.assertRaises(DataValidationError) as context:
            data_processor.verify_pallet_integrity(data)
        self.assertIn("Pallet ID did not change (still '01T26052609')", str(context.exception))
        self.assertIn("boundary marker count is 1", str(context.exception))

    def test_verify_pallet_integrity_missing_id(self):
        # Count is 1, but Pallet ID is missing
        data = [
            {"col_pallet_count": 1, "col_pallet_id": "", "_row_num": 10},
        ]
        with self.assertRaises(DataValidationError) as context:
            data_processor.verify_pallet_integrity(data)
        self.assertIn("Pallet boundary found (count=1), but Pallet ID is missing", str(context.exception))

    def test_verify_pallet_integrity_recurrence(self):
        # ID 01T26052605 reappears after 01T26052608
        data = [
            {"col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 10},
            {"col_pallet_count": 1, "col_pallet_id": "01T26052608", "_row_num": 11},
            {"col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 12},
        ]
        with self.assertRaises(DataValidationError) as context:
            data_processor.verify_pallet_integrity(data)
        self.assertIn("Pallet ID '01T26052605' reappeared after a gap", str(context.exception))

    def test_verify_pallet_integrity_empty_rows(self):
        # Empty rows should be ignored or reset the state, not crash
        data = [
            {"col_pallet_count": 1, "col_pallet_id": "01T26052605", "_row_num": 10},
            {"col_pallet_count": 0, "col_pallet_id": "", "_row_num": 11}, # completely empty
            {"col_pallet_count": 1, "col_pallet_id": "01T26052608", "_row_num": 12},
        ]
        try:
            data_processor.verify_pallet_integrity(data)
        except Exception as e:
            self.fail(f"verify_pallet_integrity failed with empty rows: {e}")

    def test_validate_no_duplicate_amount_columns(self):
        from openpyxl import Workbook
        from core.data_parser.validation import validate_no_duplicate_amount_columns, DataValidationError
        
        wb = Workbook()
        ws = wb.active
        
        # Scenario 1: No duplicates (only one Amount column)
        ws.cell(row=1, column=1, value="P.O. No.")
        ws.cell(row=1, column=2, value="Amount")
        ws.cell(row=1, column=3, value="Quantity")
        
        try:
            validate_no_duplicate_amount_columns(ws, 1)
        except DataValidationError as e:
            self.fail(f"validate_no_duplicate_amount_columns raised error unexpectedly: {e}")
            
        # Scenario 2: Duplicate col_amount columns ("Amount" and "Total value(USD)")
        ws.cell(row=1, column=4, value="Total value(USD)")
        
        with self.assertRaises(DataValidationError) as context:
            validate_no_duplicate_amount_columns(ws, 1)
            
        self.assertIn("Duplicate 'col_amount' columns detected", str(context.exception))

    def test_sheet_parser_detects_duplicate_col_amount(self):
        from openpyxl import Workbook
        from core.data_parser import sheet_parser
        from core.data_parser.validation import DataValidationError
        
        wb = Workbook()
        ws = wb.active
        
        # Row 1 has headers matching col_po, col_qty_pcs, col_amount, and another col_amount
        ws.cell(row=1, column=1, value="P.O. No.")
        ws.cell(row=1, column=2, value="Quantity")
        ws.cell(row=1, column=3, value="Amount")
        ws.cell(row=1, column=4, value="Total value(USD)")
        
        # Add some mock numeric values so the scorer qualifies the row (len(potential_mapping) >= 3, score > 0, etc.)
        ws.cell(row=2, column=1, value="PO-100")
        ws.cell(row=2, column=2, value=10)
        ws.cell(row=2, column=3, value=100.0)
        ws.cell(row=2, column=4, value=100.0)
        
        with self.assertRaises(DataValidationError) as context:
            sheet_parser.find_and_map_smart_headers(ws)
            
        self.assertIn("Duplicate 'col_amount' columns detected", str(context.exception))

    def test_validate_no_duplicate_amount_columns_pattern_matching(self):
        from openpyxl import Workbook
        from core.data_parser.validation import validate_no_duplicate_amount_columns, DataValidationError
        
        wb = Workbook()
        ws = wb.active
        
        # Column 1 has "Amount" header
        ws.cell(row=1, column=1, value="Amount")
        ws.cell(row=2, column=1, value=123.45)
        
        # Column 2 has "Price" header, value is 1.5 (< 2)
        ws.cell(row=1, column=2, value="Price")
        ws.cell(row=2, column=2, value=1.5)

        # Column 3 has empty header but data looks like amount and left neighbor is < 2
        ws.cell(row=1, column=3, value="")
        ws.cell(row=2, column=3, value=456.78)
        
        with self.assertRaises(DataValidationError) as context:
            validate_no_duplicate_amount_columns(ws, 1)
            
        self.assertIn("Duplicate 'col_amount' columns detected", str(context.exception))

    def test_validate_no_duplicate_amount_columns_adjacent_heuristic(self):
        from openpyxl import Workbook
        from core.data_parser.validation import validate_no_duplicate_amount_columns, DataValidationError
        
        wb = Workbook()
        ws = wb.active
        
        # Column 1 has "Amount" header (maps directly)
        ws.cell(row=1, column=1, value="Amount")
        ws.cell(row=2, column=1, value=100.00)
        
        # Column 2 has "Unit Price" header, value is 1.5 (< 2)
        ws.cell(row=1, column=2, value="Price")
        ws.cell(row=2, column=2, value=1.5)
        
        # Column 3 has unrecognized header (like "Duplicate Value") but value is 150.00
        # and left adjacent column value is 1.5 (< 2), so it is identified as col_amount
        ws.cell(row=1, column=3, value="Duplicate Value")
        ws.cell(row=2, column=3, value=150.00)
        
        with self.assertRaises(DataValidationError) as context:
            validate_no_duplicate_amount_columns(ws, 1)
            
        self.assertIn("Duplicate 'col_amount' columns detected", str(context.exception))

    def test_validate_no_duplicate_amount_columns_unrecognized_header_with_round_floats(self):
        from openpyxl import Workbook
        from core.data_parser.validation import validate_no_duplicate_amount_columns, DataValidationError
        
        wb = Workbook()
        ws = wb.active
        
        # Column 1 has "Amount" header (maps directly)
        ws.cell(row=1, column=1, value="Amount")
        ws.cell(row=2, column=1, value=123.00)
        
        # Column 2 has "Price" header, value is 1.5 (< 2)
        ws.cell(row=1, column=2, value="Price")
        ws.cell(row=2, column=2, value=1.5)

        # Column 3 has "StrangeName" header (unrecognized header)
        # Value is 456.00 (round float, normally stringifies to "456.0" or "456")
        ws.cell(row=1, column=3, value="StrangeName")
        ws.cell(row=2, column=3, value=456.00)
        
        with self.assertRaises(DataValidationError) as context:
            validate_no_duplicate_amount_columns(ws, 1)
            
        self.assertIn("Duplicate 'col_amount' columns detected", str(context.exception))

    def test_validate_no_duplicate_amount_columns_ignores_mapped_columns(self):
        from openpyxl import Workbook
        from core.data_parser.validation import validate_no_duplicate_amount_columns, DataValidationError
        
        wb = Workbook()
        ws = wb.active
        
        # Column 1 has "Amount" header (maps directly)
        ws.cell(row=1, column=1, value="Amount")
        ws.cell(row=2, column=1, value=100.00)
        
        # Column 2 has "Net Weight" header, value is 10.50 (which matches the amount pattern if formatted)
        ws.cell(row=1, column=2, value="Net Weight")
        ws.cell(row=2, column=2, value=10.50)
        
        # Define a column mapping where Column 2 is mapped to 'col_net'
        mapping = {
            'col_amount': 'A',
            'col_net': 'B'
        }
        
        # This should NOT raise any error because Column 2 is already mapped to 'col_net'
        try:
            validate_no_duplicate_amount_columns(ws, 1, column_mapping=mapping)
        except DataValidationError as e:
            self.fail(f"validate_no_duplicate_amount_columns raised error unexpectedly: {e}")

    def test_greedy_selection_mapping_precedence(self):
        from openpyxl import Workbook
        from core.data_parser import sheet_parser
        
        wb = Workbook()
        ws = wb.active
        
        # Column 1 has "PO" header (maps to col_po)
        ws.cell(row=1, column=1, value="P.O. No.")
        ws.cell(row=2, column=1, value="PO-100")
        
        # Column 2 has "Net Weight" header, but data is string "ABC" (weak match, score 1)
        ws.cell(row=1, column=2, value="Net Weight")
        ws.cell(row=2, column=2, value="ABC")
        
        # Column 3 has "Net Weight" header, and data is numeric 10.5 (strong match, score 5)
        ws.cell(row=1, column=3, value="Net Weight")
        ws.cell(row=2, column=3, value=10.5)
        
        # Run process_row for row 1
        mapping, score, header_text_matches = sheet_parser._process_row(ws, 1)
        
        # Column 3 (C) must be mapped to col_net because its score (5) is higher than Column 2's score (1)
        self.assertEqual(mapping.get("col_net"), "C")


if __name__ == '__main__':
    unittest.main()
