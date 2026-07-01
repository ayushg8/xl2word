from __future__ import annotations
import os
from docx import Document
from docx.shared import Emu, Pt, RGBColor
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from .model import Workbook, Sheet, Cell
from .layout import LayoutPlan, Block
from . import fit

_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT}
_DEFAULT_FONT = "Noto Sans CJK SC"   # renders Latin + Hangul; falls back if absent


def _new_document() -> Document:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = _DEFAULT_FONT
    style.font.size = Pt(10)
    # Bind an East-Asian font so CJK glyphs render.
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), _DEFAULT_FONT)
    return doc


def _section_for(doc, orientation: str):
    """Return a section in the requested orientation, starting a new section only
    when the current one's orientation differs. Keeps an all-portrait document to
    a single section while honoring per-block landscape requests."""
    section = doc.sections[-1]
    want_landscape = orientation == "landscape"
    if want_landscape == (section.orientation == WD_ORIENT.LANDSCAPE):
        return section
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    target = WD_ORIENT.LANDSCAPE if want_landscape else WD_ORIENT.PORTRAIT
    if section.orientation != target:
        section.orientation = target
        section.page_width, section.page_height = section.page_height, section.page_width
    return section


def _usable_width_emu(section) -> int:
    return int(section.page_width - section.left_margin - section.right_margin)


def _strip_extra_paragraphs(cell) -> None:
    """Remove trailing empty paragraphs left by a cell merge."""
    tc = cell._tc
    paras = tc.findall(qn("w:p"))
    while len(paras) > 1:
        last = paras[-1]
        if not "".join(t.text or "" for t in last.findall(".//" + qn("w:t"))):
            tc.remove(last)
            paras = tc.findall(qn("w:p"))
        else:
            break


def _shade(cell, rgb: str) -> None:
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), rgb)
    cell._tc.get_or_add_tcPr().append(shd)


def _fixed_layout(table) -> None:
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def _set_cell_width(cell, width_emu: int) -> None:
    cell.width = Emu(width_emu)


def _grid(sheet: Sheet, region):
    if region:
        r0, c0, r1, c1 = region
    else:
        r0, c0, r1, c1 = 1, 1, sheet.max_row, sheet.max_col
    by_pos = {(c.row, c.col): c for c in sheet.cells}
    rows = []
    for r in range(r0, r1 + 1):
        rows.append([by_pos.get((r, c)) for c in range(c0, c1 + 1)])
    return rows, (r0, c0, r1, c1)


def _apply_cell(docx_cell, model_cell: Cell | None) -> None:
    para = docx_cell.paragraphs[0]
    text = model_cell.display if model_cell else ""
    run = para.add_run(text)
    if model_cell:
        st = model_cell.style
        run.bold = st.bold
        run.italic = st.italic
        if st.font_size:
            run.font.size = Pt(st.font_size)
        if st.font_color:
            run.font.color.rgb = RGBColor.from_string(st.font_color)
        if st.align_h in _ALIGN:
            para.alignment = _ALIGN[st.align_h]
        if st.fill:
            _shade(docx_cell, st.fill)


def _pruned_axes(sheet: Sheet, rows, r0, c0, r1, c1):
    """Return the original row and column numbers to keep after dropping rows and
    columns that are fully empty within the region. A row or column that any merge
    covers is always kept, so merges stay contiguous and remap cleanly."""
    merges = [m for m in sheet.merged
              if m.min_row >= r0 and m.max_row <= r1 and m.min_col >= c0 and m.max_col <= c1]
    merge_rows, merge_cols = set(), set()
    for m in merges:
        merge_rows.update(range(m.min_row, m.max_row + 1))
        merge_cols.update(range(m.min_col, m.max_col + 1))

    def nonempty(cell):
        return bool(cell and (cell.display or "").strip())

    keep_rows = [r0 + gi for gi, row in enumerate(rows)
                 if (r0 + gi) in merge_rows or any(nonempty(c) for c in row)]
    keep_cols = [c0 + gj for gj in range(c1 - c0 + 1)
                 if (c0 + gj) in merge_cols or any(nonempty(rows[gi][gj]) for gi in range(len(rows)))]
    return keep_rows or [r0], keep_cols or [c0], merges


def _add_table(doc, sheet: Sheet, block: Block) -> None:
    full_rows, (r0, c0, r1, c1) = _grid(sheet, block.region)
    if len(full_rows) == 0 or (c1 - c0 + 1) == 0:
        return
    keep_rows, keep_cols, merges = _pruned_axes(sheet, full_rows, r0, c0, r1, c1)
    row_map = {orig: i for i, orig in enumerate(keep_rows)}
    col_map = {orig: j for j, orig in enumerate(keep_cols)}
    rows = [[full_rows[orig - r0][cc - c0] for cc in keep_cols] for orig in keep_rows]
    nrows, ncols = len(rows), len(keep_cols)

    section = _section_for(doc, block.orientation)
    text_rows = [[(cell.display if cell else "") for cell in row] for row in rows]
    usable = _usable_width_emu(section)
    widths = fit.balanced_column_widths(text_rows, usable, 10)

    table = doc.add_table(rows=nrows, cols=ncols)
    table.style = "Table Grid"
    _fixed_layout(table)
    for gi, row in enumerate(rows):
        for gj, mcell in enumerate(row):
            _apply_cell(table.cell(gi, gj), mcell)
            _set_cell_width(table.cell(gi, gj), widths[gj])
    # Push the fitted widths onto the table grid so fit-to-page actually renders.
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for gc, w_emu in zip(grid.findall(qn("w:gridCol")), widths):
            gc.set(qn("w:w"), str(int(w_emu / 635)))   # EMU -> twips
    # Repeat the first row as a header across page breaks, unless it is a title
    # banner (a lone cell merged across most of the width). Repeating a banner only
    # makes the tall title overlap the data rows on every continuation page.
    first_row = keep_rows[0]
    nonempty_first = sum(1 for c in rows[0] if c and (c.display or "").strip())
    wide_merge = any(m.min_row == first_row and (m.max_col - m.min_col + 1) >= ncols / 2
                     for m in merges)
    banner = nonempty_first <= 1 and wide_merge
    if not banner:
        _repeat_header(table.rows[0])
    # Apply merges, remapped onto the pruned grid. Every row and column a merge
    # covers is kept, so both corners are present.
    for m in merges:
        a = table.cell(row_map[m.min_row], col_map[m.min_col])
        b = table.cell(row_map[m.max_row], col_map[m.max_col])
        a.merge(b)
        _strip_extra_paragraphs(a)


def _repeat_header(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)


def _add_image(doc, images_dir: str, block: Block) -> None:
    from docx.shared import Inches
    path = block.path
    candidate = path if os.path.isabs(path) else os.path.join(os.path.dirname(images_dir), path)
    if not os.path.exists(candidate):
        candidate = os.path.join(images_dir, os.path.basename(path))
    if os.path.exists(candidate):
        doc.add_picture(candidate, width=Inches(5))
        if block.caption:
            cap = doc.add_paragraph(block.caption)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER


def write_docx(wb: Workbook, layout: LayoutPlan, out_path: str, images_dir: str) -> None:
    doc = _new_document()
    by_name = {s.name: s for s in wb.sheets}
    if layout.title:
        doc.add_heading(layout.title, level=0)
    for idx, block in enumerate(layout.blocks):
        if block.kind == "heading":
            # Put the heading in the same section orientation as the table it
            # introduces, so a landscape table does not leave its heading stranded
            # on an otherwise empty portrait page.
            for later in layout.blocks[idx + 1:]:
                if later.kind == "table":
                    _section_for(doc, later.orientation)
                    break
                if later.kind == "heading":
                    break
            doc.add_heading(block.text or "", level=block.level or 1)
        elif block.kind == "table" and block.sheet in by_name:
            _add_table(doc, by_name[block.sheet], block)
        elif block.kind == "image" and block.path:
            _add_image(doc, images_dir, block)
        elif block.kind == "pagebreak":
            doc.add_page_break()
    doc.save(out_path)
