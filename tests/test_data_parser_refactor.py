import unittest
from decimal import Decimal
from core.data_parser import data_processor
from core.data_parser import sheet_parser

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


    # --- LEGACY BEHAVIOUR VERIFICATION TESTS --- #
    def test_legacy_cbm(self):
        processed = data_processor.process_cbm_column(self.mock_legacy_extract[1].copy())
        self.assertEqual(processed['col_cbm'][0], Decimal('6.0000'))

    def test_legacy_distribute_values(self):
        processed = data_processor.distribute_values(
            self.mock_legacy_distribute[1].copy(), 
            columns_to_distribute=['col_amount'], 
            basis_column='col_qty_sf'
        )
        self.assertEqual(processed['col_amount'][0], Decimal('50.0000'))
        self.assertEqual(processed['col_amount'][1], Decimal('30.0000'))
        self.assertEqual(processed['col_amount'][2], Decimal('20.0000'))

    def test_legacy_standard_aggregation(self):
        global_map = {}
        data = self.mock_legacy_extract[1].copy()
        for k in data.copy():
            data[k].append(data[k][0])
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

if __name__ == '__main__':
    unittest.main()
