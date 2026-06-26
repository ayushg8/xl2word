import shutil, os
import pytest
from xl2word.render import render_xlsx_to_images, RenderError

soffice = shutil.which("soffice") or shutil.which("libreoffice")

@pytest.mark.skipif(not soffice, reason="LibreOffice not installed")
def test_render_xlsx_produces_png(build_simple_xlsx, tmp_path):
    pngs = render_xlsx_to_images(build_simple_xlsx(), str(tmp_path / "shots"))
    assert pngs and all(p.endswith(".png") and os.path.exists(p) for p in pngs)

def test_render_raises_clear_error_for_missing_file(tmp_path):
    with pytest.raises((RenderError, FileNotFoundError)):
        render_xlsx_to_images(str(tmp_path / "nope.xlsx"), str(tmp_path / "o"))
