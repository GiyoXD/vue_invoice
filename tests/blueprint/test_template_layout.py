import pytest
from core.blueprint_generator.internal.scanner.models import (
    TemplateLayout, UnitRow, UnitCell, TemplateMerge,
    CellStyle, FontStyle, AlignmentStyle, FillStyle, BorderStyle
)

def test_template_layout_roundtrip_new_format():
    # 1. Construct a typed TemplateLayout object
    font = FontStyle(name="Arial", size=12.0, bold=True, italic=False, color="FFFF0000")
    align = AlignmentStyle(horizontal="left", vertical="center", wrap_text=True)
    fill = FillStyle(fill_type="solid", color="FF00FF00")
    border = BorderStyle(left="thin", right="thin", top="medium", bottom="medium")
    cell_style = CellStyle(font=font, alignment=align, fill=fill, border=border, number_format="0.00")
    
    merge = TemplateMerge(min_col=1, max_col=2, row_span=1, value="Merged Header")
    
    cell1 = UnitCell(col_index=1, value="Header 1", style=cell_style, merge=merge)
    cell2 = UnitCell(col_index=2, value="Header 2", style=cell_style)
    
    row = UnitRow(relative_index=0, height=20.0, cells=[cell1, cell2])
    
    layout = TemplateLayout(
        header_rows=[row],
        footer_rows=[row]
    )
    
    # 2. Serialize to dict
    serialized = layout.to_dict()
    
    # Check JSON keys and values are correct
    assert "header_rows" in serialized
    assert "footer_rows" in serialized
    assert len(serialized["header_rows"]) == 1
    h_row = serialized["header_rows"][0]
    assert h_row["relative_index"] == 0
    assert h_row["height"] == 20.0
    assert len(h_row["cells"]) == 2
    
    c1 = h_row["cells"][0]
    assert c1["col_index"] == 1
    assert c1["value"] == "Header 1"
    assert c1["style"]["font"]["name"] == "Arial"
    assert c1["merge"]["min_col"] == 1
    assert c1["merge"]["max_col"] == 2
    assert c1["merge"]["value"] == "Merged Header"
    
    # 3. Deserialize from dict
    deserialized = TemplateLayout.from_dict(serialized)
    
    # 4. Verify deserialized layout matches original
    assert len(deserialized.header_rows) == 1
    d_row = deserialized.header_rows[0]
    assert d_row.relative_index == 0
    assert d_row.height == 20.0
    assert len(d_row.cells) == 2
    assert d_row.cells[0].col_index == 1
    assert d_row.cells[0].value == "Header 1"
    assert d_row.cells[0].style.font.name == "Arial"
    assert d_row.cells[0].merge.value == "Merged Header"
    
    assert len(deserialized.footer_rows) == 1
    assert deserialized.footer_rows[0].relative_index == 0
    assert deserialized.footer_rows[0].height == 20.0
    
    assert deserialized.header_rows[0].cells[0].style.font.name == "Arial"
    assert deserialized.header_rows[0].cells[0].style.alignment.wrap_text is True


def test_translate_template_rows_merge_translation():
    from core.invoice_generator.builders.json_template_builder import translate_template_rows
    
    # Template columns:
    # Col 1: normal, Col 2: merge start (spans 2-4), Col 3: inside merge, Col 4: inside merge
    cell1 = UnitCell(col_index=1, value="Normal")
    cell2 = UnitCell(
        col_index=2,
        value="Merged",
        merge=TemplateMerge(min_col=2, max_col=4, row_span=1, value="Merged")
    )
    cell3 = UnitCell(col_index=3, value=None)
    cell4 = UnitCell(col_index=4, value=None)
    
    row = UnitRow(relative_index=0, height=15.0, cells=[cell1, cell2, cell3, cell4])
    
    # Case A: Column 2 (merge start) is hidden, but 3 and 4 are visible.
    # Mapping: {1: 1, 2: None, 3: 2, 4: 3}
    mapping = {1: 1, 2: None, 3: 2, 4: 3}
    translated = translate_template_rows([row], mapping)
    
    assert len(translated) == 1
    t_cells = translated[0].cells
    
    # cell1 is mapped to 1
    # cell2 (originally col 2) is shifted to 2 (first visible column of the merge)
    # cell3 and cell4 are covered by the template merge range and should be skipped
    assert len(t_cells) == 2
    assert t_cells[0].col_index == 1
    assert t_cells[0].value == "Normal"
    
    assert t_cells[1].col_index == 2
    assert t_cells[1].value == "Merged"
    assert t_cells[1].merge is not None
    assert t_cells[1].merge.min_col == 2
    assert t_cells[1].merge.max_col == 3  # last visible column of the merge
    assert t_cells[1].merge.row_span == 1
    assert t_cells[1].merge.value == "Merged"


def test_translate_template_rows_all_hidden_merge():
    from core.invoice_generator.builders.json_template_builder import translate_template_rows
    
    cell1 = UnitCell(
        col_index=2,
        value="Merged",
        merge=TemplateMerge(min_col=2, max_col=3, row_span=1, value="Merged")
    )
    row = UnitRow(relative_index=0, height=15.0, cells=[cell1])
    
    # Case: all columns in merge are hidden
    mapping = {2: None, 3: None}
    translated = translate_template_rows([row], mapping)
    
    assert len(translated) == 1
    # Cell is skipped entirely since no columns in the merge are visible
    assert len(translated[0].cells) == 0



