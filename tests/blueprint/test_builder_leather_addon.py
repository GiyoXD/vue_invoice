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
    col_pallet_id = ColumnInfo(id="col_pallet_id", header="Pallet ID", col_index=2, width=10.0)
    col_desc = ColumnInfo(id="col_desc", header="Description", col_index=3, width=10.0)
    col_qty_pcs = ColumnInfo(id="col_qty_pcs", header="Qty Pcs", col_index=4, width=10.0)
    col_net = ColumnInfo(id="col_net", header="Net Weight", col_index=5, width=10.0)
    sheet.columns = [col_po, col_pallet_id, col_desc, col_qty_pcs, col_net]

    # Setup static content hints containing the detected leather summaries
    from core.blueprint_generator.internal.scanner.models.addons import LeatherSummaryFact
    # Setup static content hints containing the pre-built addon facts
    sheet.static_content_hints = {
        "addon_facts": [
            LeatherSummaryFact(
                leather_key="BUFFALO",
                total_col_id="col_po",
                label_col_id="col_pallet_id",
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
        pallet_count_col_id="col_pallet_id"
    )

    builder = ConfigBuilder()
    footer_data = builder._build_footer(sheet)
    rows = footer_data["rows"]

    # Total rows: 1 main footer row + 1 leather row = 2 rows
    assert len(rows) == 2
    
    # Verify Buffalo row structure
    buffalo_row = rows[1]
    assert any(c["col_id"] == "col_po" and c["value"] == "TOTAL OF:" for c in buffalo_row)
    assert any(c["col_id"] == "col_pallet_id" and c["value"] == "BUFFALO LEATHER" for c in buffalo_row)
    assert any(c["col_id"] == "col_desc" and c["value"] == "{pallet_count} PALLET{multiple}" for c in buffalo_row)
    # col_qty_pcs and col_net should be dynamically appended because they exist in sheet_col_ids
    assert any(c["col_id"] == "col_qty_pcs" and "value" not in c for c in buffalo_row)
    assert any(c["col_id"] == "col_net" and "value" not in c for c in buffalo_row)


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
    rows = footer_data["rows"]

    # Since the pattern was not found, we expect only the Main footer row.
    assert len(rows) == 1
