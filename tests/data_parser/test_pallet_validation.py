import decimal
import unittest
from decimal import Decimal
from unittest.mock import MagicMock, patch

from core.data_parser.validation import (
    DataValidationError,
    validate_data,
    verify_pallet_integrity,
)


class TestPalletPairingPhase1(unittest.TestCase):
    def test_missing_pallet_id_with_sqft_logs_warning(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_id": "",
                "col_qty_sf": Decimal("100"),
            }
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_called_once()
        warning_msg = monitor.log_warning.call_args[0][0]
        self.assertIn(
            "[Pallet Pairing] 1 row(s) have sqft data but missing Pallet ID: rows 1. Operator may have forgotten to fill these.",
            warning_msg,
        )

    def test_zero_sqft_missing_id_no_warning(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_id": "",
                "col_qty_sf": Decimal("0"),
            }
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_not_called()

    def test_missing_pallet_id_truncation_over_10_rows(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": i,
                "col_pallet_id": "",
                "col_qty_sf": Decimal("10"),
            }
            for i in range(1, 13)
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_called_once()
        warning_msg = monitor.log_warning.call_args[0][0]
        self.assertIn("(and 2 more)", warning_msg)

    def test_missing_pallet_id_column_skips_with_warning(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_count": 1,
                "col_qty_sf": Decimal("100"),
            }
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        self.assertTrue(monitor.log_warning.called)
        warning_msgs = [call[0][0] for call in monitor.log_warning.call_args_list]
        self.assertTrue(
            any("col_pallet_id column not found" in msg for msg in warning_msgs)
        )

    def test_formatted_sqft_strings_parsed(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_id": "",
                "col_qty_sf": "1,250.50",
            }
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_called_once()
        warning_msg = monitor.log_warning.call_args[0][0]
        self.assertIn(
            "[Pallet Pairing] 1 row(s) have sqft data but missing Pallet ID: rows 1. Operator may have forgotten to fill these.",
            warning_msg,
        )


class TestPalletCorrelationPhase2(unittest.TestCase):
    def test_valid_single_and_multi_row_pallets(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "P01"},
            {"_row_num": 2, "col_pallet_count": 1, "col_pallet_id": "P02"},
            {"_row_num": 3, "col_pallet_count": 0, "col_pallet_id": "P02"},
            {"_row_num": 4, "col_pallet_count": 0, "col_pallet_id": "P02"},
            {"_row_num": 5, "col_pallet_count": 1, "col_pallet_id": "P03"},
        ]
        verify_pallet_integrity(rows)

    def test_pallet_id_missing_when_count_is_one(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": ""},
        ]
        with self.assertRaises(DataValidationError) as ctx:
            verify_pallet_integrity(rows)
        self.assertIn("Pallet ID missing (count = 1)", str(ctx.exception))

    def test_pallet_id_changed_but_count_is_zero(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "P01"},
            {"_row_num": 2, "col_pallet_count": 0, "col_pallet_id": "P02"},
        ]
        with self.assertRaises(DataValidationError) as ctx:
            verify_pallet_integrity(rows)
        self.assertIn("Pallet ID changed to 'P02' but count is 0", str(ctx.exception))

    def test_pallet_id_same_but_count_is_one(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "P01"},
            {"_row_num": 2, "col_pallet_count": 1, "col_pallet_id": "P01"},
        ]
        with self.assertRaises(DataValidationError) as ctx:
            verify_pallet_integrity(rows)
        self.assertIn(
            "Pallet count is 1 but Pallet ID did not change", str(ctx.exception)
        )

    def test_pallet_id_reappeared_after_gap(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "P01"},
            {"_row_num": 2, "col_pallet_count": 1, "col_pallet_id": "P02"},
            {"_row_num": 3, "col_pallet_count": 1, "col_pallet_id": "P01"},
        ]
        with self.assertRaises(DataValidationError) as ctx:
            verify_pallet_integrity(rows)
        self.assertIn("reappeared after gap", str(ctx.exception))

    def test_spacer_rows_between_pallets(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "P01"},
            {"_row_num": 2, "col_pallet_count": 0, "col_pallet_id": ""},
            {"_row_num": 3, "col_pallet_count": 1, "col_pallet_id": "P02"},
        ]
        verify_pallet_integrity(rows)

    def test_spacer_row_continuation(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "P01"},
            {"_row_num": 2, "col_pallet_count": 0, "col_pallet_id": ""},
            {"_row_num": 3, "col_pallet_count": 0, "col_pallet_id": "P01"},
        ]
        verify_pallet_integrity(rows)

    def test_whitespace_stripped(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": " P01 "},
            {"_row_num": 2, "col_pallet_count": 0, "col_pallet_id": "P01\t"},
            {"_row_num": 3, "col_pallet_count": 1, "col_pallet_id": "  P02  "},
        ]
        verify_pallet_integrity(rows)


class TestPalletAsymmetricHeaders(unittest.TestCase):
    def test_pallet_id_exists_without_count(self):
        monitor = MagicMock()
        rows = [
            {"_row_num": 1, "col_pallet_id": "", "col_qty_sf": Decimal("100")},
            {"_row_num": 2, "col_pallet_id": "P02", "col_qty_sf": Decimal("100")},
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_called_once()
        self.assertIn(
            "[Pallet Pairing] 1 row(s) have sqft data but missing Pallet ID: rows 1",
            monitor.log_warning.call_args[0][0],
        )

    def test_pallet_count_exists_without_id(self):
        monitor = MagicMock()
        rows = [
            {"_row_num": 1, "col_pallet_count": 1},
            {"_row_num": 2, "col_pallet_count": 0, "col_cbm": Decimal("2.5")},
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        warning_msgs = [call[0][0] for call in monitor.log_warning.call_args_list]
        self.assertTrue(
            any("col_pallet_id column not found" in msg for msg in warning_msgs)
        )
        self.assertTrue(
            any("CBM (2.5) is on a sub-row" in msg for msg in warning_msgs)
        )

    def test_neither_pallet_column_exists(self):
        monitor = MagicMock()
        rows = [
            {"_row_num": 1, "col_item": "ITEM-1", "col_qty_sf": Decimal("100")},
            {"_row_num": 2, "col_item": "ITEM-2", "col_qty_sf": Decimal("200")},
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_not_called()


class TestPalletMetricAlignmentPhase3(unittest.TestCase):
    def test_sub_row_with_metrics_logs_warnings(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_count": 1,
                "col_pallet_id": "P01",
                "col_cbm": Decimal("1.2"),
                "col_net": Decimal("1000"),
                "col_gross": Decimal("1100"),
            },
            {
                "_row_num": 2,
                "col_pallet_count": 0,
                "col_pallet_id": "P01",
                "col_cbm": Decimal("0.5"),
                "col_net": Decimal("500"),
                "col_gross": Decimal("550"),
            },
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        self.assertEqual(monitor.log_warning.call_count, 3)
        warning_msgs = [call[0][0] for call in monitor.log_warning.call_args_list]
        self.assertTrue(any("CBM" in msg for msg in warning_msgs))
        self.assertTrue(any("Net Weight" in msg for msg in warning_msgs))
        self.assertTrue(any("Gross Weight" in msg for msg in warning_msgs))

    def test_pallet_label_in_sub_row_warning(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_count": 1,
                "col_pallet_id": "P01",
            },
            {
                "_row_num": 2,
                "col_pallet_count": 0,
                "col_pallet_id": "P01",
                "col_cbm": Decimal("0.8"),
            },
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_called_once()
        warning_msg = monitor.log_warning.call_args[0][0]
        self.assertIn("(Pallet 'P01')", warning_msg)


class TestPalletValidationIntegration(unittest.TestCase):
    def test_validate_data_dispatcher_runs_pallet_integrity(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_count": 1,
                "col_pallet_id": "",
            }
        ]
        with self.assertRaises(DataValidationError) as ctx:
            validate_data(
                data_rows=rows,
                table_id_str="Table 1",
                column_mapping={"col_pallet_id": "A", "col_pallet_count": "B"},
                monitor=monitor,
                phase="integrity",
            )
        self.assertIn("Pallet ID missing (count = 1)", str(ctx.exception))

    def test_validate_data_presence_phase_does_not_run_pallet(self):
        monitor = MagicMock()
        with patch(
            "core.data_parser.validation.dispatcher.verify_pallet_integrity"
        ) as mock_pallet, patch(
            "core.data_parser.validation.dispatcher.validate_table_data_presence"
        ) as mock_presence:
            validate_data(
                data_rows=[],
                table_id_str="Table 1",
                column_mapping={},
                monitor=monitor,
                phase="presence",
            )
            mock_presence.assert_called_once()
            mock_pallet.assert_not_called()


class TestPalletBugFixesAndEdgeCases(unittest.TestCase):
    def test_case_insensitive_pallet_id_continuity(self):
        rows = [
            {"_row_num": 1, "col_pallet_count": 1, "col_pallet_id": "01t26052605"},
            {"_row_num": 2, "col_pallet_count": 0, "col_pallet_id": "01T26052605"},
        ]
        verify_pallet_integrity(rows)

    def test_pallet_count_range_strings(self):
        for range_val in ["1-19", "1/19", "1 of 19", "1 PLT"]:
            rows = [
                {"_row_num": 1, "col_pallet_count": range_val, "col_pallet_id": "P01"},
                {"_row_num": 2, "col_pallet_count": 0, "col_pallet_id": "P01"},
            ]
            verify_pallet_integrity(rows)

    def test_spacer_row_without_pallet_id_no_metric_warning(self):
        monitor = MagicMock()
        rows = [
            {
                "_row_num": 1,
                "col_pallet_count": 1,
                "col_pallet_id": "P01",
                "col_cbm": Decimal("1.0"),
            },
            {
                "_row_num": 2,
                "col_pallet_count": 0,
                "col_pallet_id": "",
                "col_cbm": Decimal("0.5"),
            },
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        monitor.log_warning.assert_not_called()

    def test_pallet_id_not_hijacked_by_pallet_no(self):
        monitor = MagicMock()
        rows = [
            {"_row_num": 1, "col_pallet_no": "P01", "col_pallet_count": 1},
            {"_row_num": 2, "col_pallet_no": "P01", "col_pallet_count": 0},
            {"_row_num": 3, "col_pallet_no": "P02", "col_pallet_count": 1},
        ]
        verify_pallet_integrity(rows, monitor=monitor)
        self.assertTrue(monitor.log_warning.called)
        warning_msgs = [call[0][0] for call in monitor.log_warning.call_args_list]
        self.assertTrue(
            any("col_pallet_id column not found" in msg for msg in warning_msgs)
        )

    def test_sheet_parser_mercy_rule_and_case_insensitivity(self):
        from core.data_parser.sheet_parser import _score_cell

        self.assertEqual(_score_cell("col_pallet_id", None, 1, 1), 1)
        self.assertEqual(_score_cell("col_pallet_id", "", 1, 1), 1)
        self.assertEqual(_score_cell("col_pallet_id", "01t26052605", 1, 1), 15)
        self.assertEqual(_score_cell("col_pallet_count", 0, 1, 1), 0)
        self.assertEqual(_score_cell("col_pallet_count", 1, 1, 1), 25)


if __name__ == "__main__":
    unittest.main()
