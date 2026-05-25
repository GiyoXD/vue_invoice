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
        self.assertIn("Data Validation Error (Row 12)", str(context.exception))
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
        self.assertIn("Data Validation Error (Rows 14 and 15)", str(context.exception))
        self.assertIn("more pieces (100) but a lower CBM (1.0)", str(context.exception))

        # With ignore_cbm_warning=True, it should pass
        try:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion', ignore_cbm_warning=True)
        except DataValidationError:
            self.fail("validate_data raised DataValidationError when ignore_cbm_warning was True")

    def test_cbm_pcs_proportion_abnormally_high_ratio(self):
        # ITEM1 has col_qty_pcs=2 and col_cbm=1.5 -> ratio = 0.75 > 0.5
        data = [
            {"col_po": "PO1", "col_item": "ITEM1", "col_qty_pcs": Decimal('2'), "col_cbm": Decimal('1.5'), "_row_num": 16}
        ]
        column_mapping = {"col_po": "A", "col_item": "B", "col_qty_pcs": "C", "col_cbm": "D"}

        with self.assertRaises(DataValidationError) as context:
            validate_data(data, "Table 1", column_mapping, phase='cbm_proportion')
        self.assertIn("Data Validation Error (Row 16)", str(context.exception))
        self.assertIn("abnormally high CBM-to-basis ratio", str(context.exception))

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
        self.assertIn("Data Validation Error (Row 21)", str(context.exception))
        self.assertIn("outlier (< 0.1x median ratio", str(context.exception))

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

if __name__ == '__main__':
    unittest.main()
