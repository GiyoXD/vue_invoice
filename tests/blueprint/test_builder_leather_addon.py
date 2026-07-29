import pytest
from unittest.mock import MagicMock
from core.blueprint_generator.internal.builder import ConfigBuilder
from core.blueprint_generator.internal.scanner import SheetAnalysis
from core.blueprint_generator.internal.scanner.models import (
    TemplateLayout, UnitRow, UnitCell, ColumnInfo, FooterInfo
)

def test_config_builder_leather_summary_detection():
    # 1. Create a mock sheet that has the "total of:" and "leather" pattern
    sheet = MagicMock(spec=SheetAnalysis)
    sheet.name = "Packing list"
    sheet.data_source = "processed_tables_multi"
    
    # Columns: col_po (col 1), col_pallet_id (col 2), col_desc (col 3), col_qty_pcs (col 4), col_net (col 5)
    col_po = ColumnInfo(id="col_po", header="PO", col_index=1, width=10.0)
    col_pallet_no = ColumnInfo(id="col_pallet_no", header="Pallet ID", col_index=2, width=10.0)
    col_desc = ColumnInfo(id="col_desc", header="Description", col_index=3, width=10.0)
    col_qty_pcs = ColumnInfo(id="col_qty_pcs", header="Qty Pcs", col_index=4, width=10.0)
    col_net = ColumnInfo(id="col_net", header="Net Weight", col_index=5, width=10.0)
    sheet.columns = [col_po, col_pallet_no, col_desc, col_qty_pcs, col_net]

    # Setup static content hints containing the detected leather summaries
    from core.blueprint_generator.internal.scanner.models.addons import LeatherSummaryFact
    # Setup static content hints containing the pre-built addon facts
    sheet.static_content_hints = {
        "addon_facts": [
            LeatherSummaryFact(
                leather_key="BUFFALO",
                total_col_id="col_po",
                label_col_id="col_pallet_no",
                total_value="TOTAL OF:",
                label_value="BUFFALO LEATHER"
            )
        ]
    }
    sheet.footer_info = FooterInfo(
        row_num=10,
        total_text="TOTAL:",
        total_text_col_id="col_po",
        merge_curr_colspan=1,
        pallet_count_col_id="col_pallet_no"
    )

    builder = ConfigBuilder()
    footer_data = builder._build_footer(sheet)
    summary_data = builder._build_summary(sheet)
    footer_rows = footer_data["rows"]
    summary_rows = summary_data["rows"]

    # Footer total row: 1 main row
    assert len(footer_rows) == 1
    # Summary rows: 0 weight summary rows + 1 leather row = 1 row
    assert len(summary_rows) == 1
    
    # Verify Buffalo row structure
    buffalo_row = summary_rows[0]
    assert isinstance(buffalo_row, dict)
    assert buffalo_row["source_list"] == "leather_summary"
    cells = buffalo_row["cells"]
    assert any(c["col_id"] == "col_po" and c.get("value") == "TOTAL OF:" for c in cells)
    assert any(c["col_id"] == "col_pallet_no" and c.get("value") == "{leather_type} LEATHER" for c in cells)
    assert any(c["col_id"] == "col_desc" and c.get("value") == "{col_pallet_count} PALLET{multiple}" for c in cells)
    # col_qty_pcs and col_net should be dynamically appended because they exist in sheet_col_ids
    assert any(c["col_id"] == "col_qty_pcs" for c in cells)
    assert any(c["col_id"] == "col_net" for c in cells)


def test_config_builder_leather_summary_skipped_when_no_pattern():
    # 2. Create a mock sheet that does NOT have the pattern
    sheet = MagicMock(spec=SheetAnalysis)
    sheet.name = "Packing list"
    sheet.data_source = "processed_tables_multi"
    
    col_po = ColumnInfo(id="col_po", header="PO", col_index=1, width=10.0)
    sheet.columns = [col_po]

    # Setup static content hints without any leather summaries
    sheet.static_content_hints = {}
    sheet.footer_info = FooterInfo(
        row_num=10,
        total_text="TOTAL:",
        total_text_col_id="col_po",
        merge_curr_colspan=1
    )

    builder = ConfigBuilder()
    footer_data = builder._build_footer(sheet)
    summary_data = builder._build_summary(sheet)
    footer_rows = footer_data["rows"]
    summary_rows = summary_data["rows"]

    # Main footer row = 1, weight summary rows = 0
    assert len(footer_rows) == 1
    assert len(summary_rows) == 0


