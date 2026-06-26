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


def _add_table(doc, sheet: Sheet, block: Block) -> None:
    rows, (r0, c0, r1, c1) = _grid(sheet, block.region)
    nrows, ncols = len(rows), (c1 - c0 + 1)
    if nrows == 0 or ncols == 0:
        return
    section = _section_for(doc, block.orientation)
    text_rows = [[(cell.display if cell else "") for cell in row] for row in rows]
    natural = fit.natural_column_widths(text_rows, 10)
    usable = _usable_width_emu(section)
    widths = fit.fit_columns(natural, usable)

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
    # Repeat header row across page breaks.
    _repeat_header(table.rows[0])
    # Apply merges that fall inside this region.
    for m in sheet.merged:
        if m.min_row >= r0 and m.max_row <= r1 and m.min_col >= c0 and m.max_col <= c1:
            a = table.cell(m.min_row - r0, m.min_col - c0)
            b = table.cell(m.max_row - r0, m.max_col - c0)
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
    for block in layout.blocks:
        if block.kind == "heading":
            doc.add_heading(block.text or "", level=block.level or 1)
        elif block.kind == "table" and block.sheet in by_name:
            _add_table(doc, by_name[block.sheet], block)
        elif block.kind == "image" and block.path:
            _add_image(doc, images_dir, block)
        elif block.kind == "pagebreak":
            doc.add_page_break()
    doc.save(out_path)
