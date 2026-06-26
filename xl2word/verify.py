from __future__ import annotations
from docx import Document
from .render import render_docx_to_images


def render_doc(docx_path: str, out_dir: str) -> list[str]:
    return render_docx_to_images(docx_path, out_dir)


def detect_overflow(docx_path: str) -> list[str]:
    doc = Document(docx_path)
    section = doc.sections[0]
    usable = int(section.page_width - section.left_margin - section.right_margin)
    issues: list[str] = []
    for i, table in enumerate(doc.tables):
        # Read true per-column widths from the table grid. A merged banner cell
        # reports the full span width on every grid cell it covers, so summing row
        # cells overcounts; the grid carries one width per column. Widen a column
        # only for a non-spanning cell that genuinely exceeds its grid width.
        col_widths = [int(c.width) if c.width is not None else 0 for c in table.columns]
        for row in table.rows:
            for j, cell in enumerate(row.cells):
                if j >= len(col_widths) or (cell._tc.grid_span or 1) > 1:
                    continue
                w = cell.width
                if w is not None and int(w) > col_widths[j]:
                    col_widths[j] = int(w)
        total = sum(col_widths)
        if total > usable + 2:  # 2 EMU tolerance
            issues.append(
                f"Table {i} is wider than the page "
                f"({total / 914400:.2f}in > {usable / 914400:.2f}in usable)."
            )
    return issues
