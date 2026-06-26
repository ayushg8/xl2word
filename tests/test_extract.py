from xl2word.extract import extract_semantic

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
