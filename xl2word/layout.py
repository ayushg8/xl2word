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


def segment_regions(sheet: Sheet) -> list[tuple]:
    """Split a sheet into its natural content blocks: maximal rectangles of
    populated cells separated by empty rows and columns. A sheet often stacks
    several differently shaped tables (a narrow pre-set block over a wide recipe,
    two material lists side by side). Rendering each block as its own right-sized
    table keeps every column wide enough instead of squeezing them all to the
    width of the sheet's widest block. Rows or columns a merge spans are never used
    as a separator, so merged headers stay intact."""
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

    def col_spanned(cc, band_rows):  # a merge intersecting the band crosses this column
        return any(b < cc <= e and b < e and not (c < band_rows[0] or a > band_rows[1])
                   for (a, b, c, e) in merges)

    occ_rows = {r for r, _ in full}
    r_lo, r_hi = min(occ_rows), max(occ_rows)
    bands, cur = [], []
    for r in range(r_lo, r_hi + 1):
        if r in occ_rows:
            cur.append(r)
        elif cur and not row_spanned(r):
            bands.append(cur); cur = []
        elif cur:
            cur.append(r)
    if cur:
        bands.append(cur)

    regions = []
    for band in bands:
        br0, br1 = band[0], band[-1]
        cset = {c for (r, c) in full if br0 <= r <= br1}
        groups, g = [], []
        for c in range(min(cset), max(cset) + 1):
            if c in cset:
                g.append(c)
            elif g and not col_spanned(c, (br0, br1)):
                groups.append(g); g = []
            elif g:
                g.append(c)
        if g:
            groups.append(g)
        for grp in groups:
            sub = [(r, c) for (r, c) in full if br0 <= r <= br1 and grp[0] <= c <= grp[-1]]
            rr = [r for r, _ in sub]; cc = [c for _, c in sub]
            regions.append((min(rr), min(cc), max(rr), max(cc)))
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
