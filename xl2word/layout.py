from __future__ import annotations
import json
from dataclasses import dataclass, field, asdict
from .model import Workbook, Sheet
from . import fit

_EMU = 914400
_PORTRAIT_USABLE = _EMU * 65 // 10   # 6.5in on letter with 1in margins
_LANDSCAPE_USABLE = _EMU * 9         # 9.0in on letter landscape


@dataclass
class Block:
    kind: str
    text: str | None = None
    level: int | None = None
    sheet: str | None = None
    region: tuple | None = None
    orientation: str = "portrait"
    path: str | None = None
    caption: str | None = None


@dataclass
class LayoutPlan:
    title: str
    blocks: list[Block] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2)

    @classmethod
    def from_json(cls, s: str) -> "LayoutPlan":
        d = json.loads(s)
        blocks = []
        for b in d["blocks"]:
            region = tuple(b["region"]) if b.get("region") else None
            blocks.append(Block(
                kind=b["kind"], text=b.get("text"), level=b.get("level"),
                sheet=b.get("sheet"), region=region,
                orientation=b.get("orientation", "portrait"),
                path=b.get("path"), caption=b.get("caption"),
            ))
        return cls(title=d["title"], blocks=blocks)


_ANCHOR_MIN_ROWS = 4   # a side-by-side gap must span at least this many rows to cut on


def segment_regions(sheet: Sheet) -> list[tuple]:
    """Split a sheet into its natural content blocks by recursive X-Y cutting.

    A sheet often stacks several differently shaped tables (a narrow pre-set block
    over a wide recipe, two material lists side by side). Rendering each block as
    its own right-sized table keeps every column wide enough instead of squeezing
    them all to the width of the sheet's widest block.

    Each region is cut on the first available of: a fully-empty row, a fully-empty
    column, or an "anchored gap" -- a column empty for a run of rows anchored to the
    region's top or bottom with content on both sides, which is how a full-width
    table butts directly against two side-by-side blocks with no separating blank
    row. Rows and columns a merge spans are never used as a cut, so merged headers
    stay intact."""
    disp = {(c.row, c.col): (c.display or "").strip() for c in sheet.cells}
    merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in sheet.merged]
    full = {k for k, v in disp.items() if v}
    for (r0, c0, r1, c1) in merges:
        if disp.get((r0, c0)):
            full.update((r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1))
    if not full:
        return []

    def row_spanned(r):   # a merge crosses the boundary just above row r
        return any(a < r <= c and a < c for (a, b, c, e) in merges)

    def col_spanned(cc, rlo, rhi):  # a merge intersecting [rlo,rhi] crosses left of cc
        return any(b < cc <= e and b < e and not (c < rlo or a > rhi)
                   for (a, b, c, e) in merges)

    def has(rlo, rhi, clo, chi):
        return clo <= chi and any((r, c) in full
                                  for r in range(rlo, rhi + 1) for c in range(clo, chi + 1))

    def bbox(rlo, clo, rhi, chi):
        pts = [(r, c) for (r, c) in full if rlo <= r <= rhi and clo <= c <= chi]
        if not pts:
            return None
        return (min(r for r, _ in pts), min(c for _, c in pts),
                max(r for r, _ in pts), max(c for _, c in pts))

    def two_pass(rlo, clo, rhi, chi):
        """One level: split into row bands, then column groups within each band.
        Deliberately not recursive, so a sparse column (scattered answers) stays a
        single block instead of shattering into one cell per value."""
        region = bbox(rlo, clo, rhi, chi)
        if not region:
            return []
        rlo, clo, rhi, chi = region
        rowset = {r for (r, c) in full if rlo <= r <= rhi and clo <= c <= chi}
        bands, s = [], None
        for r in range(rlo, rhi + 1):
            if r in rowset:
                s = [r, r] if s is None else [s[0], r]
            elif s and not row_spanned(r):
                bands.append(tuple(s)); s = None
            elif s:
                s = [s[0], r]
        if s:
            bands.append(tuple(s))
        out = []
        for br0, br1 in bands:
            colset = {c for (r, c) in full if br0 <= r <= br1 and clo <= c <= chi}
            groups, g = [], None
            for c in range(min(colset), max(colset) + 1):
                if c in colset:
                    g = [c, c] if g is None else [g[0], c]
                elif g and not col_spanned(c, br0, br1):
                    groups.append(tuple(g)); g = None
                elif g:
                    g = [g[0], c]
            if g:
                groups.append(tuple(g))
            for gc0, gc1 in groups:
                b = bbox(br0, gc0, br1, gc1)
                if b:
                    out.append(b)
        return out

    def anchored_cut(rlo, clo, rhi, chi):
        """Row at which an interior column's top- or bottom-anchored empty run ends,
        with content on both sides of it -- a full-width table butting against two
        side-by-side blocks. None if no such gap. Only for regions wide enough to
        plausibly hold side-by-side blocks, so narrow log columns are left alone."""
        if chi - clo + 1 < 6:
            return None
        best = None
        for cc in range(clo + 1, chi):
            top = next((r for r in range(rlo, rhi + 1) if (r, cc) in full), rhi + 1) - rlo
            if top >= _ANCHOR_MIN_ROWS and not row_spanned(rlo + top) \
                    and has(rlo, rlo + top - 1, clo, cc - 1) and has(rlo, rlo + top - 1, cc + 1, chi):
                if best is None or top > best[0]:
                    best = (top, rlo + top)
            bot = rhi - next((r for r in range(rhi, rlo - 1, -1) if (r, cc) in full), rlo - 1)
            if bot >= _ANCHOR_MIN_ROWS and not row_spanned(rhi - bot + 1) \
                    and has(rhi - bot + 1, rhi, clo, cc - 1) and has(rhi - bot + 1, rhi, cc + 1, chi):
                if best is None or bot > best[0]:
                    best = (bot, rhi - bot + 1)
        return best[1] if best else None

    def refine(reg, depth):
        rlo, clo, rhi, chi = reg
        k = anchored_cut(rlo, clo, rhi, chi) if depth < 3 else None
        if k is None:
            return [reg]
        out = []
        for half in ((rlo, clo, k - 1, chi), (k, clo, rhi, chi)):
            for sub in two_pass(*half):
                out.extend(refine(sub, depth + 1))
        return out

    r0 = min(r for r, _ in full); r1 = max(r for r, _ in full)
    c0 = min(c for _, c in full); c1 = max(c for _, c in full)
    regions = []
    for reg in two_pass(r0, c0, r1, c1):
        regions.extend(refine(reg, 0))
    return regions


def _region_orientation(sheet: Sheet, region) -> str:
    r0, c0, r1, c1 = region
    by = {(c.row, c.col): c for c in sheet.cells}
    rows = [[(by.get((r, c)).display if by.get((r, c)) else "") for c in range(c0, c1 + 1)]
            for r in range(r0, r1 + 1)]
    total = sum(fit.balanced_column_widths(rows, _LANDSCAPE_USABLE, 10))
    return "portrait" if total <= _PORTRAIT_USABLE else "landscape"


def _is_banner(sheet: Sheet, region) -> bool:
    """A lone cell merged across most of the region's width: a title, not a table."""
    r0, c0, r1, c1 = region
    ncols = c1 - c0 + 1
    filled = [c for c in sheet.cells
              if r0 <= c.row <= r1 and c0 <= c.col <= c1 and (c.display or "").strip()]
    if len(filled) != 1:
        return False
    return any(m.min_row <= filled[0].row <= m.max_row
               and (m.max_col - m.min_col + 1) >= ncols / 2 for m in sheet.merged)


def default_layout(wb: Workbook) -> LayoutPlan:
    import os
    blocks: list[Block] = []
    for i, sheet in enumerate(wb.sheets):
        if i > 0:
            blocks.append(Block(kind="pagebreak"))
        blocks.append(Block(kind="heading", text=sheet.name, level=1))
        for region in segment_regions(sheet):
            if _is_banner(sheet, region):
                cell = next(c for c in sheet.cells
                            if c.row == region[0] and (c.display or "").strip())
                blocks.append(Block(kind="heading", text=cell.display, level=2))
            else:
                blocks.append(Block(kind="table", sheet=sheet.name, region=region,
                                    orientation=_region_orientation(sheet, region)))
        for img in sheet.images:
            blocks.append(Block(kind="image", path=img.path,
                                caption=os.path.basename(img.path)))
    # Workbook-level media not anchored to a sheet: append after the first sheet.
    if not any(s.images for s in wb.sheets):
        for img in wb.media:
            blocks.append(Block(kind="image", path=img.path,
                                caption=os.path.basename(img.path)))
    title = os.path.splitext(os.path.basename(wb.source))[0]
    return LayoutPlan(title=title, blocks=blocks)
