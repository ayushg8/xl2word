import json, os
from xl2word.extract import extract_semantic, extract_media, extract_workbook

def test_semantic_captures_values_styles_merges(build_rich_xlsx):
    wb = extract_semantic(build_rich_xlsx())
    s = wb.sheets[0]
    assert s.name == "Cell Design"
    cells = {(c.row, c.col): c for c in s.cells}
    assert cells[(1, 1)].display == "ESS LFP Cell Design"
    assert cells[(1, 1)].style.bold is True
    assert cells[(1, 1)].style.fill == "FFD966"
    assert cells[(3, 2)].display == "1.2%"            # 0.012 with 0.0%
    assert any(m.min_row == 1 and m.max_col == 3 for m in s.merged)

def test_simple_col_width(build_simple_xlsx):
    wb = extract_semantic(build_simple_xlsx())
    assert wb.sheets[0].col_widths.get(1) == 18

def test_media_sweep_catches_embedded_image(build_rich_xlsx, tmp_path):
    imgs = extract_media(build_rich_xlsx(), str(tmp_path / "imgs"))
    assert len(imgs) >= 1
    assert os.path.exists(tmp_path / "imgs" / os.path.basename(imgs[0].path))

def test_extract_workbook_writes_json(build_simple_xlsx, tmp_path):
    out = str(tmp_path / "out")
    wb = extract_workbook(build_simple_xlsx(), out, render=False)
    assert os.path.exists(os.path.join(out, "workbook.json"))
    data = json.load(open(os.path.join(out, "workbook.json")))
    assert data["sheets"][0]["name"] == "Spec"
