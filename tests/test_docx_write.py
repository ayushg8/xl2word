import os
from docx import Document
from xl2word.extract import extract_semantic
from xl2word.layout import default_layout
from xl2word.docx_write import write_docx

def test_write_docx_builds_editable_table(build_simple_xlsx, tmp_path):
    wb = extract_semantic(build_simple_xlsx())
    out = str(tmp_path / "out.docx")
    write_docx(wb, default_layout(wb), out, images_dir=str(tmp_path / "images"))
    assert os.path.exists(out)
    doc = Document(out)
    assert len(doc.tables) == 1
    t = doc.tables[0]
    flat = [c.text for row in t.rows for c in row.cells]
    assert "Loading Level" in flat
    assert "25.40" in flat            # display value preserved
    assert "Parameter" in flat

def test_merged_header_is_merged(build_rich_xlsx, tmp_path):
    wb = extract_semantic(build_rich_xlsx(with_image=False))
    out = str(tmp_path / "rich.docx")
    write_docx(wb, default_layout(wb), out, images_dir=str(tmp_path / "images"))
    doc = Document(out)
    row0 = doc.tables[0].rows[0]
    # A1:C1 merged -> the three grid cells resolve to one underlying cell
    assert row0.cells[0]._tc is row0.cells[2]._tc
