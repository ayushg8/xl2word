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


def choose_orientation(natural_total_emu: int, portrait_usable_emu: int,
                       landscape_usable_emu: int) -> str:
    if natural_total_emu > portrait_usable_emu and \
       natural_total_emu <= landscape_usable_emu * 1.0:
        return "landscape"
    if natural_total_emu > portrait_usable_emu:
        return "landscape"
    return "portrait"
