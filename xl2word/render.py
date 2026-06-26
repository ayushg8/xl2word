from __future__ import annotations
import os
import shutil
import subprocess


class RenderError(RuntimeError):
    pass


def _soffice() -> str:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise RenderError("LibreOffice (soffice) not found on PATH; needed for rendering.")
    return exe


def _soffice_to_pdf(src: str, out_dir: str) -> str:
    if not os.path.exists(src):
        raise FileNotFoundError(src)
    os.makedirs(out_dir, exist_ok=True)
    subprocess.run(
        [_soffice(), "--headless", "--convert-to", "pdf", "--outdir", out_dir, src],
        check=True, capture_output=True, timeout=120,
    )
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RenderError(f"soffice did not produce a PDF for {src}")
    return pdf


def _pdf_to_pngs(pdf_path: str, out_dir: str, dpi: int) -> list[str]:
    import fitz  # PyMuPDF
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    paths = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for i, page in enumerate(doc):
        png = os.path.join(out_dir, f"{stem}_p{i + 1}.png")
        page.get_pixmap(matrix=mat).save(png)
        paths.append(png)
    doc.close()
    return paths


def render_xlsx_to_images(xlsx_path: str, out_dir: str, dpi: int = 150) -> list[str]:
    pdf = _soffice_to_pdf(xlsx_path, out_dir)
    return _pdf_to_pngs(pdf, out_dir, dpi)


def render_docx_to_images(docx_path: str, out_dir: str, dpi: int = 150) -> list[str]:
    pdf = _soffice_to_pdf(docx_path, out_dir)
    return _pdf_to_pngs(pdf, out_dir, dpi)
