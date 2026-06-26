from xl2word.fit import (EMU_PER_INCH, natural_column_widths, fit_columns,
                         choose_orientation)

def test_natural_widths_track_content_length():
    w = natural_column_widths([["a", "wider content here"], ["b", "x"]], 10)
    assert w[1] > w[0]

def test_fit_scales_down_when_over_budget():
    natural = [5 * EMU_PER_INCH, 5 * EMU_PER_INCH]  # 10in natural
    fitted = fit_columns(natural, 6 * EMU_PER_INCH)  # 6in usable
    assert abs(sum(fitted) - 6 * EMU_PER_INCH) <= 2
    assert fitted[0] == fitted[1]

def test_fit_keeps_natural_when_under_budget():
    natural = [2 * EMU_PER_INCH, 2 * EMU_PER_INCH]
    assert fit_columns(natural, 6 * EMU_PER_INCH) == natural

def test_orientation_flips_for_wide():
    assert choose_orientation(9 * EMU_PER_INCH, 6 * EMU_PER_INCH, 9 * EMU_PER_INCH) == "landscape"
    assert choose_orientation(5 * EMU_PER_INCH, 6 * EMU_PER_INCH, 9 * EMU_PER_INCH) == "portrait"
