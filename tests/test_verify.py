from docx import Document
from docx.shared import Inches
from xl2word.verify import detect_overflow

def test_detect_overflow_flags_wide_table(tmp_path):
    doc = Document()
    section = doc.sections[0]
    usable_in = (section.page_width - section.left_margin - section.right_margin) / 914400
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).width = Inches(usable_in)         # each col = full usable width
    t.cell(0, 1).width = Inches(usable_in)         # total = 2x usable -> overflow
    p = str(tmp_path / "wide.docx")
    doc.save(p)
    issues = detect_overflow(p)
    assert any("wider than" in s.lower() for s in issues)

def test_no_overflow_when_narrow(tmp_path):
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).width = Inches(1)
    t.cell(0, 1).width = Inches(1)
    p = str(tmp_path / "narrow.docx")
    doc.save(p)
    assert detect_overflow(p) == []
