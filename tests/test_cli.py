import os
from docx import Document
from xl2word.cli import main

def test_cli_converts_xlsx_to_docx(build_simple_xlsx, tmp_path):
    src = build_simple_xlsx()
    out = str(tmp_path / "out.docx")
    rc = main([src, "-o", out, "--workdir", str(tmp_path / "wd"), "--no-render"])
    assert rc == 0
    assert os.path.exists(out)
    assert len(Document(out).tables) == 1
