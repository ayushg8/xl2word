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
_DENSE_FRAC = 0.35     # a column counts as populated if this fraction of content rows fill it
_VGAP_ROWS = 3         # a contiguous empty-row run this tall splits a column group vertically


def _is_substantial(gc0, gc1, br0, br1, full) -> bool:
    """A column group is a real side-by-side table (worth splitting off) only if it
    has at least two well-populated columns. A group with one dense column padded by
    near-empty ones is a sparse annotation strip (e.g. an answer/status column) that
    belongs with its neighbour, not a standalone table -- splitting it off renders a
    broken, mostly-blank fragment."""
    content_rows = [r for r in range(br0, br1 + 1)
                    if any((r, c) in full for c in range(gc0, gc1 + 1))]
    if not content_rows:
        return False
    dense = sum(1 for c in range(gc0, gc1 + 1)
                if sum(1 for r in content_rows if (r, c) in full) >= len(content_rows) * _DENSE_FRAC)
    return dense >= 2


def _absorb_thin_groups(groups, br0, br1, full):
    """Attach a thin annotation column-group to the substantial table on its LEFT,
    so answer/status columns rejoin the questions they belong to (which read left of
    them). A thin block with no table to its left -- a standalone note list sitting
    left of an unrelated table -- is kept as its own region rather than being pulled
    into that table and interleaved. If no group is substantial, the band is one
    table."""
    if len(groups) <= 1:
        return groups
    subs = [_is_substantial(a, b, br0, br1, full) for (a, b) in groups]
    if not any(subs):
        return [(groups[0][0], groups[-1][1])]
    spans = {}
    for i, ok in enumerate(subs):
        if ok:
            spans[i] = list(groups[i])
    for i, grp in enumerate(groups):
        if subs[i]:
            continue
        left_anchors = [j for j in spans if groups[j][0] < grp[0]]
        if left_anchors:
            j = max(left_anchors, key=lambda k: groups[k][0])   # nearest table on the left
            spans[j][1] = max(spans[j][1], grp[1])
        else:
            spans[i] = list(grp)                                # standalone; keep separate
    return [tuple(spans[i]) for i in sorted(spans, key=lambda k: spans[k][0])]


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
            groups = _absorb_thin_groups(groups, br0, br1, full)
            for gc0, gc1 in groups:
                # Re-split this column group at genuine vertical gaps (a contiguous
                # run of >=3 empty rows), which a band-level scan misses when other
                # columns carry content across the gap. Stacked blocks in one column
                # (a params table above a maintenance-note list) become separate
                # tables instead of one grid with blank rows and misaligned columns.
                grp_rows = [r for r in range(br0, br1 + 1)
                            if any((r, c) in full for c in range(gc0, gc1 + 1))]
                start = prev = None
                for r in grp_rows:
                    if start is None:
                        start = prev = r
                    elif r - prev - 1 >= _VGAP_ROWS and not any(row_spanned(x) for x in range(prev + 1, r + 1)):
                        b = bbox(start, gc0, prev, gc1)
                        if b:
                            out.append(b)
                        start = prev = r
                    else:
                        prev = r
                if start is not None:
                    b = bbox(start, gc0, prev, gc1)
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
    """Portrait unless the table genuinely cannot fit portrait width. Judge by the
    minimum width each column needs (its longest word), not the roomy natural width,
    so a modest table stays portrait instead of jumping to a half-empty landscape
    page; only tables whose columns cannot be squeezed into portrait go landscape."""
    r0, c0, r1, c1 = region
    by = {(c.row, c.col): c for c in sheet.cells}
    rows = [[(by.get((r, c)).display if by.get((r, c)) else "") for c in range(c0, c1 + 1)]
            for r in range(r0, r1 + 1)]
    if fit.natural_width_sum(rows, 10) <= _PORTRAIT_USABLE:
        return "portrait"                                  # fits portrait comfortably
    if fit.min_width_sum(rows, 10) <= _PORTRAIT_USABLE * 72 // 100:
        return "portrait"                                  # squeezes into portrait with room to spare
    return "landscape"                                     # genuinely needs the wider page


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


def _norm(s: str) -> str:
    return "".join(ch for ch in s.lower() if ch.isalnum())


def _has_body(sheet: Sheet, region) -> bool:
    """Skip a region that is just a wide column-label strip with an empty body (an
    unfilled template table -- prints as a near-blank grid). Keep everything with
    data in two or more rows, and keep a genuine tiny table (a lone row of one or
    two cells is real content, not a header)."""
    r0, c0, r1, c1 = region
    per_row = {}
    for c in sheet.cells:
        if r0 <= c.row <= r1 and c0 <= c.col <= c1 and (c.display or "").strip():
            per_row[c.row] = per_row.get(c.row, 0) + 1
    if len(per_row) >= 2:
        return True
    if not per_row:
        return False
    return next(iter(per_row.values())) < 3   # lone row: keep only if it is not a label strip


def _promote_banners(sheet: Sheet, region):
    """Split a region at its internal section-banner rows (a lone cell merged
    across most of the block's width, e.g. "Pre-set Conditions" sitting above its
    own little table). Returns an ordered list of ("heading", text) and
    ("table", subregion) items so those banners become real headings that read as
    sections and show up in the contents, instead of merged rows buried in a table."""
    r0, c0, r1, c1 = region
    ncols = c1 - c0 + 1
    cells = {(c.row, c.col): c for c in sheet.cells if (c.display or "").strip()}
    disp = {k: v.display.strip() for k, v in cells.items()}
    banners = {}
    for r in range(r0, r1 + 1):
        filled = [c for c in range(c0, c1 + 1) if (r, c) in cells]
        if len(filled) != 1:
            continue
        cell = cells[(r, filled[0])]
        # A section title is bold; a long merged note is not a heading. Allow a
        # long title only on the region's first row (a document/section title,
        # e.g. "IBC Cell Spec Matrix — ..."), never a long mid-table note.
        if not cell.style.bold or (len(disp[(r, filled[0])]) > 60 and r != r0):
            continue
        width = max((m.max_col - m.min_col + 1 for m in sheet.merged
                     if m.min_row <= r <= m.max_row and m.min_col <= filled[0] <= m.max_col),
                    default=1)
        if width >= max(2, ncols * 0.8):
            banners[r] = disp[(r, filled[0])]
    if not banners:
        return [("table", region)]

    items, start = [], None
    for r in range(r0, r1 + 1):
        if r in banners:
            if start is not None:
                items.append(("table", (start, c0, r - 1, c1))); start = None
            items.append(("heading", banners[r]))
        elif start is None:
            start = r
    if start is not None:
        items.append(("table", (start, c0, r1, c1)))

    out = []
    for kind, val in items:
        if kind == "heading":
            out.append((kind, val))
        elif any(disp.get((r, c)) for r in range(val[0], val[2] + 1)
                 for c in range(val[1], val[3] + 1)):
            out.append((kind, val))
    return out


def default_layout(wb: Workbook) -> LayoutPlan:
    import os
    blocks: list[Block] = []
    content_sheets = [s for s in wb.sheets if s.cells]   # skip empty sheets entirely
    for i, sheet in enumerate(content_sheets):
        # Sheets flow continuously rather than each forced onto a fresh page: a
        # forced break strands a section's short table tail on a near-empty page.
        # The bold H1 heading (kept with its content) delineates each sheet.
        blocks.append(Block(kind="heading", text=sheet.name, level=1))
        sheet_key = _norm(sheet.name)

        def add_section_heading(text):
            # Skip a banner that just restates the sheet name (e.g. sheet
            # "Material Spec" with banner "MATERIAL SPEC - IQC"). Match on prefix so
            # a distinct title that merely contains the sheet name is kept.
            key = _norm(text)
            if sheet_key and (key.startswith(sheet_key) or sheet_key.startswith(key)):
                return
            blocks.append(Block(kind="heading", text=text, level=2))

        for region in segment_regions(sheet):
            if _is_banner(sheet, region):
                cell = next(c for c in sheet.cells
                            if c.row == region[0] and (c.display or "").strip())
                add_section_heading(cell.display)
                continue
            for kind, val in _promote_banners(sheet, region):
                if kind == "heading":
                    add_section_heading(val)
                elif _has_body(sheet, val):
                    blocks.append(Block(kind="table", sheet=sheet.name, region=val,
                                        orientation=_region_orientation(sheet, val)))
        for img in sheet.images:
            blocks.append(Block(kind="image", path=img.path,
                                caption=os.path.basename(img.path)))
    # Workbook-level media not anchored to a sheet: append after the first sheet.
    if not any(s.images for s in wb.sheets):
        for img in wb.media:
            blocks.append(Block(kind="image", path=img.path,
                                caption=os.path.basename(img.path)))
    return LayoutPlan(title=_doc_title(wb, blocks), blocks=blocks)


_GENERIC_NAMES = {"input", "output", "book1", "book", "sheet1", "workbook", "untitled", "temp", "tmp"}


def _doc_title(wb: Workbook, blocks: list[Block]) -> str:
    """Prefer the workbook filename as the cover title. If it is a generic export
    name like 'input.xlsx', fall back to the first sheet's own title banner (the
    document's real title, e.g. 'ASSEMBLY LINE OCAP ...') — only from the first
    sheet, so we don't grab an unrelated section title deeper in the doc."""
    import os
    name = os.path.splitext(os.path.basename(wb.source))[0].strip()
    if name and name.lower() not in _GENERIC_NAMES:
        return name
    # Generic export name: use the first sheet that has content (its name is a
    # reliable title; a section banner deeper in the doc is not).
    for s in wb.sheets:
        if s.cells:
            return s.name
    return name or "Document"
