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
        widths = []
        for cell in table.rows[0].cells:
            w = cell.width
            widths.append(int(w) if w is not None else 0)
        total = sum(widths)
        if total > usable + 2:  # 2 EMU tolerance
            issues.append(
                f"Table {i} is wider than the page "
                f"({total / 914400:.2f}in > {usable / 914400:.2f}in usable)."
            )
    return issues
