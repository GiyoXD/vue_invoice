import pytest
from openpyxl import Workbook
from core.invoice_generator.builders.table.builder import TableBuilder
from core.invoice_generator.models.layout import SheetLayoutState
from core.invoice_generator.models.config.layout import SheetLayoutModel
from core.invoice_generator.models.config.styling import SheetStylingModel
from core.invoice_generator.models.table_adapter import ResolvedTableData

def test_builder_skip_all_builders():
    wb = Workbook()
    ws = wb.active
    
    sheet_config = {
        "structure": {
            "columns": [
              {"id": "col_a", "colspan": 1}
            ]
        }
    }
    
    layout_state = SheetLayoutState()
    sheet_layout = SheetLayoutModel.model_validate({"structure": {"columns": sheet_config["structure"]["columns"]}})
    sheet_styling = SheetStylingModel()
    resolved_data = ResolvedTableData()
    
    builder = TableBuilder(
        workbook=wb,
        worksheet=ws,
        sheet_styling=sheet_styling,
        sheet_layout=sheet_layout,
        resolved_data=resolved_data,
        sheet_name="Test Sheet",
        skip_header_builder=True,
        skip_data_table_builder=True,
        skip_footer_builder=True,
        layout_state=layout_state
    )
    
    # Run orchestration
    success = builder.build(start_row=5)
    
    assert success is True
    # The build succeeded but skipped all sub-builders
    # The grid advanced row only by the header skip fallback (2 rows)
    assert builder.grid._cursor_row == 2
    assert builder.next_row_after_footer == 7
