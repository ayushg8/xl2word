from __future__ import annotations

EMU_PER_INCH = 914400
_MIN_COL_EMU = EMU_PER_INCH // 2  # 0.5in floor so columns never collapse


def estimate_text_width_emu(text: str, font_size_pt: float) -> int:
    # Average glyph advance ~0.5em; pad one em for cell margins.
    longest = max((len(line) for line in str(text).split("\n")), default=0)
    em_emu = font_size_pt * EMU_PER_INCH / 72.0
    return int((longest * 0.5 + 1.0) * em_emu)


def natural_column_widths(rows: list[list[str]], font_size_pt: float) -> list[int]:
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    widths = [_MIN_COL_EMU] * ncols
    for r in rows:
        for i, val in enumerate(r):
            widths[i] = max(widths[i], estimate_text_width_emu(val, font_size_pt))
    return widths


def fit_columns(natural: list[int], usable_width_emu: int) -> list[int]:
    total = sum(natural)
    if total <= usable_width_emu or total == 0:
        return list(natural)
    scale = usable_width_emu / total
    return [max(1, int(w * scale)) for w in natural]


def _text_width_emu(chars: float, font_size_pt: float) -> int:
    # 0.55em average glyph advance, plus ~1.6em padding to cover Word's default
    # 0.08in per-side cell margins so short labels do not wrap a character early.
    em_emu = font_size_pt * EMU_PER_INCH / 72.0
    return int((chars * 0.55 + 1.6) * em_emu)


def balanced_column_widths(rows: list[list[str]], usable_width_emu: int,
                           font_size_pt: float) -> list[int]:
    """Fit columns to the page without starving short label columns.

    Uniform proportional scaling collapses a 7-character label to two letters
    whenever one neighbor holds a paragraph of free text. Instead: give every
    column at least the width of its longest single word (so labels never wrap
    mid-word), cap any one column so a long text field wraps rather than dominates,
    and when the total still overflows, take the excess only from columns that have
    slack above their minimum."""
    if not rows:
        return []
    ncols = max((len(r) for r in rows), default=0)
    floor = EMU_PER_INCH // 2               # 0.5in absolute minimum
    cap = EMU_PER_INCH * 26 // 10           # 2.6in cap for any single column
    natural = [floor] * ncols
    minw = [floor] * ncols
    for r in rows:
        for i, val in enumerate(r):
            lines = str(val).split("\n")
            longest_line = max((len(ln) for ln in lines), default=0)
            longest_word = max((len(w) for ln in lines for w in ln.split(" ")), default=0)
            natural[i] = max(natural[i], _text_width_emu(longest_line, font_size_pt))
            minw[i] = max(minw[i], _text_width_emu(longest_word, font_size_pt))
    target = [min(max(n, m), max(cap, m)) for n, m in zip(natural, minw)]
    total = sum(target)
    if total <= usable_width_emu:
        return target
    excess = total - usable_width_emu
    slack = sum(t - m for t, m in zip(target, minw))
    if slack >= excess and slack > 0:
        return [t - (t - m) * excess // slack for t, m in zip(target, minw)]
    # Even the minimum word widths overflow: scale the minimums to fit.
    scale = usable_width_emu / sum(minw)
    return [max(1, int(m * scale)) for m in minw]


def choose_orientation(natural_total_emu: int, portrait_usable_emu: int,
                       landscape_usable_emu: int) -> str:
    if natural_total_emu > portrait_usable_emu and \
       natural_total_emu <= landscape_usable_emu * 1.0:
        return "landscape"
    if natural_total_emu > portrait_usable_emu:
        return "landscape"
    return "portrait"
