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
        footer_rows=[row],
        col_widths={"A": 15.0, "B": 20.0}
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
    
    assert serialized["col_widths"]["A"] == 15.0
    
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
    
    assert deserialized.col_widths["A"] == 15.0
    assert deserialized.header_rows[0].cells[0].style.font.name == "Arial"
    assert deserialized.header_rows[0].cells[0].style.alignment.wrap_text is True


