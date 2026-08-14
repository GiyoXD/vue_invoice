import pytest
from core.invoice_generator.utils.formula_formatter import extract_decimal_places, wrap_with_round
from core.invoice_generator.builders.table.table_grid import TableGrid
from core.invoice_generator.styling.style_registry import StyleRegistry

def test_extract_decimal_places():
    assert extract_decimal_places(None) is None
    assert extract_decimal_places("") is None
    assert extract_decimal_places("@") is None
    assert extract_decimal_places("General") is None
    assert extract_decimal_places("Text") is None
    assert extract_decimal_places("0.00") == 2
    assert extract_decimal_places("#,##0.000") == 3
    assert extract_decimal_places("#,##0.0000") == 4
    assert extract_decimal_places("$#,##0.00;($#,##0.00);\"-\"") == 2
    assert extract_decimal_places("0.0%") == 1
    assert extract_decimal_places("0") == 0
    assert extract_decimal_places("#,##0") == 0

def test_wrap_with_round():
    assert wrap_with_round("A1*B1", None) == "A1*B1"
    assert wrap_with_round("A1*B1", 2) == "=ROUND(A1*B1, 2)"
    assert wrap_with_round("=A1*B1", 2) == "=ROUND(A1*B1, 2)"
    assert wrap_with_round("=ROUND(A1*B1, 2)", 2) == "=ROUND(A1*B1, 2)"
    assert wrap_with_round("ROUND(A1*B1, 2)", 2) == "=ROUND(A1*B1, 2)"
    assert wrap_with_round("", 2) == ""
    assert wrap_with_round("=SUM(A1:A10)", 0) == "=ROUND(SUM(A1:A10), 0)"

def test_write_formula():
    column_mapping = {"col_qty": 1, "col_price": 2, "col_amount": 3}
    style_registry = StyleRegistry({
        "columns": {
            "col_qty": {"format": "0"},
            "col_price": {"format": "0.00"},
            "col_amount": {"format": "#,##0.00"},
        },
        "row_contexts": {}
    })
    grid = TableGrid(column_mapping, style_registry)
    grid.set_start_row(1)

    # Write formula for col_amount with decimal format -> should wrap with ROUND(..., 2)
    grid.write_formula(0, "col_amount", "={col_ref_0}*{col_ref_1}", ["col_qty", "col_price"])
    cell = grid.get_cell(0, "col_amount")
    assert cell.value == "=ROUND(A1*B1, 2)"

    # Write formula for col_qty with format "0" -> should wrap with ROUND(..., 0)
    grid.write_formula(0, "col_qty", "={col_ref_0}*2", ["col_price"])
    cell_qty = grid.get_cell(0, "col_qty")
    assert cell_qty.value == "=ROUND(B1*2, 0)"

    # Grid without formatting for a column
    style_registry_no_fmt = StyleRegistry({
        "columns": {
            "col_no_fmt": {"format": "@"}
        },
        "row_contexts": {}
    })
    grid_no_fmt = TableGrid({"col_no_fmt": 1, "col_price": 2}, style_registry_no_fmt)
    grid_no_fmt.set_start_row(1)
    grid_no_fmt.write_formula(0, "col_no_fmt", "={col_ref_0}*10", ["col_price"])
    cell_no_fmt = grid_no_fmt.get_cell(0, "col_no_fmt")
    assert cell_no_fmt.value == "=B1*10"

def test_write_section_aggregate():
    column_mapping = {"col_amount": 3}
    style_registry = StyleRegistry({
        "columns": {
            "col_amount": {"format": "#,##0.0000"}
        },
        "row_contexts": {}
    })
    grid = TableGrid(column_mapping, style_registry)
    grid.set_start_row(1)
    grid.set_section_bounds("data", 0, 9)

    grid.write_section_aggregate(10, "col_amount", function="SUM", section="data")
    cell = grid.get_cell(10, "col_amount")
    assert cell.value == "=SUM(C1:C10)"

def test_write_formula_and_section_aggregate_no_style_registry():
    column_mapping = {"col_qty": 1, "col_price": 2, "col_amount": 3}
    grid = TableGrid(column_mapping, style_registry=None)
    grid.set_start_row(1)
    grid.set_section_bounds("data", 0, 9)

    # write_formula with style_registry=None
    grid.write_formula(0, "col_amount", "={col_ref_0}*{col_ref_1}", ["col_qty", "col_price"])
    cell_formula = grid.get_cell(0, "col_amount")
    assert cell_formula.value == "=A1*B1"

    # write_section_aggregate with style_registry=None
    grid.write_section_aggregate(10, "col_amount", function="SUM", section="data")
    cell_agg = grid.get_cell(10, "col_amount")
    assert cell_agg.value == "=SUM(C1:C10)"

