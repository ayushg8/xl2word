# xl2word Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a general Excel→Word converter — any `.xlsx` in, a clean, editable, neatly-fitted `.docx` out — with a deterministic Python capture core and a Claude-driven layout/verify skill on top.

**Architecture:** Python deterministically extracts everything from the workbook (semantic cells + raw-zip media + per-sheet screenshots) into a JSON model. A deterministic renderer turns a structured `LayoutPlan` into an editable `.docx` with fitted tables. The `SKILL.md` orchestration adds the Claude stages: understand the data, author `layout.md`/`layout.json`, render, then screenshot-verify the output against the contract.

**Tech Stack:** Python 3.11+, openpyxl (read), python-docx (write), PyMuPDF (PDF→PNG), Pillow (image dims), LibreOffice/`soffice` (xlsx/docx→PDF render), pytest.

## Global Constraints

- Python **3.11+** (uses `X | None` syntax, `tomllib` not required).
- Dependencies, exact: `openpyxl`, `python-docx`, `pymupdf`, `pillow`, `pytest`. System: LibreOffice providing `soffice` on PATH (rendering only).
- Rows and columns are **1-based** everywhere (matches openpyxl).
- RGB colors are stored as 6-char `'RRGGBB'` strings; unknown/theme colors are `None`.
- Tables in the output are **native, editable Word tables** — never images — except the explicit per-region screenshot fallback.
- **No AI tells in any generated document text**: no em dashes, no "comprehensive/robust/leverage/delve/underscore/crucial", no filler. Generated titles/captions are plain.
- Every module is import-safe without LibreOffice; only `render.py` functions require `soffice` at call time.

## File Structure

```
IBC/xl2word/
  README.md                 # Task 13
  pyproject.toml            # Task 1
  .gitignore                # Task 1
  SPEC.md                   # exists
  PLAN.md                   # this file
  SKILL.md                  # Task 12 — the Claude orchestration skill
  xl2word/
    __init__.py             # Task 1
    model.py                # Task 2 — Workbook/Sheet/Cell/Style/MergedRange/ImageAsset + JSON
    cleaners.py             # Task 3 — format_cell_value(value, number_format)
    extract.py              # Task 4,5 — extract_semantic / extract_media / extract_workbook
    render.py               # Task 6 — xlsx|docx -> PNGs via soffice + PyMuPDF
    fit.py                  # Task 7 — column width + orientation math
    layout.py               # Task 8 — Block/LayoutPlan + default_layout
    docx_write.py           # Task 9 — write_docx(wb, layout, out_path, images_dir)
    verify.py               # Task 10 — render_doc + detect_overflow
    cli.py                  # Task 11 — argparse entry point
  tests/
    conftest.py             # Task 2 — fixture builders (build_simple_xlsx, etc.)
    test_model.py           # Task 2
    test_cleaners.py        # Task 3
    test_extract.py         # Task 4,5
    test_render.py          # Task 6
    test_fit.py             # Task 7
    test_layout.py          # Task 8
    test_docx_write.py      # Task 9
    test_verify.py          # Task 10
    test_cli.py             # Task 11
    test_integration.py     # Task 13
  examples/                 # Task 13
```

**Two entry paths.** The **CLI** runs extract → `default_layout` → `write_docx` (deterministic, scriptable). The **skill** runs the same Python pieces but inserts Claude's understand/plan/verify stages and authors a richer `layout.json`. `write_docx` consumes a structured `LayoutPlan`; Claude also writes a human-readable `layout.md` mirror used for screenshot verification.

### Core interfaces (the consistent spine — every task uses these exact names)

```python
# model.py
@dataclass
class Style:
    bold: bool = False
    italic: bool = False
    font_name: str | None = None
    font_size: float | None = None
    font_color: str | None = None      # 'RRGGBB'
    fill: str | None = None            # 'RRGGBB' background, None = no fill
    align_h: str | None = None         # 'left'|'center'|'right'
    align_v: str | None = None         # 'top'|'center'|'bottom'
    border: bool = False               # any visible border on the cell
    number_format: str | None = None

@dataclass
class Cell:
    row: int; col: int                 # 1-based
    value: object                      # raw cell value (may be None)
    display: str                       # formatted text, '' when value is None
    style: Style
    hyperlink: str | None = None
    note: str | None = None

@dataclass
class MergedRange:
    min_row: int; min_col: int; max_row: int; max_col: int

@dataclass
class ImageAsset:
    id: str
    path: str                          # path relative to the out_dir images folder
    sheet: str | None = None
    anchor_row: int | None = None
    anchor_col: int | None = None
    width_px: int | None = None
    height_px: int | None = None
    source: str = 'media'              # 'openpyxl' | 'media'

@dataclass
class Sheet:
    name: str; index: int
    max_row: int; max_col: int
    cells: list[Cell]
    merged: list[MergedRange]
    col_widths: dict[int, float]       # col -> Excel width units
    images: list[ImageAsset]
    screenshots: list[str]             # rel paths

@dataclass
class Workbook:
    source: str
    sheets: list[Sheet]
    media: list[ImageAsset]
    has_charts: bool = False
    has_embeddings: bool = False
    def to_json(self) -> str: ...
    @classmethod
    def from_json(cls, s: str) -> "Workbook": ...

# cleaners.py
def format_cell_value(value, number_format: str | None) -> str

# extract.py
def extract_semantic(xlsx_path: str) -> Workbook
def extract_media(xlsx_path: str, images_dir: str) -> list[ImageAsset]
def extract_workbook(xlsx_path: str, out_dir: str, render: bool = True) -> Workbook

# render.py
class RenderError(RuntimeError): ...
def render_xlsx_to_images(xlsx_path: str, out_dir: str, dpi: int = 150) -> list[str]
def render_docx_to_images(docx_path: str, out_dir: str, dpi: int = 150) -> list[str]

# fit.py
EMU_PER_INCH = 914400
def estimate_text_width_emu(text: str, font_size_pt: float) -> int
def natural_column_widths(rows: list[list[str]], font_size_pt: float) -> list[int]
def fit_columns(natural: list[int], usable_width_emu: int) -> list[int]
def choose_orientation(natural_total_emu: int, portrait_usable_emu: int, landscape_usable_emu: int) -> str

# layout.py
@dataclass
class Block:
    kind: str                          # 'heading'|'table'|'image'|'pagebreak'
    text: str | None = None
    level: int | None = None
    sheet: str | None = None
    region: tuple | None = None        # (min_row,min_col,max_row,max_col) or None = whole sheet
    orientation: str = 'portrait'
    path: str | None = None
    caption: str | None = None
@dataclass
class LayoutPlan:
    title: str
    blocks: list[Block]
    def to_json(self) -> str: ...
    @classmethod
    def from_json(cls, s: str) -> "LayoutPlan": ...
def default_layout(wb: Workbook) -> LayoutPlan

# docx_write.py
def write_docx(wb: Workbook, layout: LayoutPlan, out_path: str, images_dir: str) -> None

# verify.py
def render_doc(docx_path: str, out_dir: str) -> list[str]
def detect_overflow(docx_path: str) -> list[str]

# cli.py
def main(argv: list[str] | None = None) -> int
```

---

### Task 1: Repo scaffold + git init

**Files:**
- Create: `xl2word/pyproject.toml`, `xl2word/.gitignore`, `xl2word/xl2word/__init__.py`, `xl2word/tests/__init__.py`

**Interfaces:**
- Produces: an installable package `xl2word` and a green (empty) pytest run.

- [ ] **Step 1: Create `pyproject.toml`**

```toml
[project]
name = "xl2word"
version = "0.1.0"
description = "General Excel to Word converter: deterministic capture, Claude-driven layout."
requires-python = ">=3.11"
dependencies = ["openpyxl>=3.1", "python-docx>=1.1", "pymupdf>=1.24", "pillow>=10.0"]

[project.optional-dependencies]
dev = ["pytest>=8.0"]

[project.scripts]
xl2word = "xl2word.cli:main"

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[tool.setuptools.packages.find]
include = ["xl2word*"]
```

- [ ] **Step 2: Create `.gitignore`**

```gitignore
__pycache__/
*.pyc
.pytest_cache/
*.egg-info/
build/
dist/
.DS_Store
/scratch/
*.pdf
```

- [ ] **Step 3: Create empty package files**

`xl2word/xl2word/__init__.py`:
```python
"""xl2word: general Excel to Word converter."""
__version__ = "0.1.0"
```
`xl2word/tests/__init__.py`: (empty file)

- [ ] **Step 4: Install and verify pytest runs green-empty**

Run: `cd xl2word && pip install -e ".[dev]" && pytest -q`
Expected: `no tests ran` (exit 5) or `0 passed` — no import errors.

- [ ] **Step 5: git init + commit**

```bash
cd xl2word && git init && git add -A
git commit -m "chore: scaffold xl2word package"
```

---

### Task 2: `model.py` data model + JSON round-trip + fixture builders

**Files:**
- Create: `xl2word/xl2word/model.py`, `xl2word/tests/conftest.py`, `xl2word/tests/test_model.py`

**Interfaces:**
- Produces: all dataclasses in the spine (`Style`, `Cell`, `MergedRange`, `ImageAsset`, `Sheet`, `Workbook`), `Workbook.to_json/from_json`, and pytest fixtures `build_simple_xlsx`, `build_rich_xlsx` used by later tasks.

- [ ] **Step 1: Write the failing test** — `tests/test_model.py`

```python
from xl2word.model import Workbook, Sheet, Cell, Style, MergedRange, ImageAsset

def test_workbook_json_roundtrip():
    wb = Workbook(
        source="x.xlsx",
        sheets=[Sheet(
            name="S1", index=0, max_row=1, max_col=2,
            cells=[Cell(1, 1, "Loading", "Loading", Style(bold=True)),
                   Cell(1, 2, 25.4, "25.40", Style(number_format="0.00"))],
            merged=[MergedRange(1, 1, 1, 2)],
            col_widths={1: 12.0, 2: 8.0},
            images=[ImageAsset(id="img1", path="images/img1.png", sheet="S1")],
            screenshots=["screenshots/S1.png"],
        )],
        media=[ImageAsset(id="img1", path="images/img1.png")],
    )
    restored = Workbook.from_json(wb.to_json())
    assert restored == wb
    assert restored.sheets[0].cells[1].display == "25.40"
    assert restored.sheets[0].merged[0].max_col == 2
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_model.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xl2word.model'`

- [ ] **Step 3: Write `model.py`**

```python
from __future__ import annotations
import json
from dataclasses import dataclass, field, asdict


@dataclass
class Style:
    bold: bool = False
    italic: bool = False
    font_name: str | None = None
    font_size: float | None = None
    font_color: str | None = None
    fill: str | None = None
    align_h: str | None = None
    align_v: str | None = None
    border: bool = False
    number_format: str | None = None


@dataclass
class Cell:
    row: int
    col: int
    value: object
    display: str
    style: Style
    hyperlink: str | None = None
    note: str | None = None


@dataclass
class MergedRange:
    min_row: int
    min_col: int
    max_row: int
    max_col: int


@dataclass
class ImageAsset:
    id: str
    path: str
    sheet: str | None = None
    anchor_row: int | None = None
    anchor_col: int | None = None
    width_px: int | None = None
    height_px: int | None = None
    source: str = "media"


@dataclass
class Sheet:
    name: str
    index: int
    max_row: int
    max_col: int
    cells: list[Cell] = field(default_factory=list)
    merged: list[MergedRange] = field(default_factory=list)
    col_widths: dict[int, float] = field(default_factory=dict)
    images: list[ImageAsset] = field(default_factory=list)
    screenshots: list[str] = field(default_factory=list)


@dataclass
class Workbook:
    source: str
    sheets: list[Sheet] = field(default_factory=list)
    media: list[ImageAsset] = field(default_factory=list)
    has_charts: bool = False
    has_embeddings: bool = False

    def to_json(self) -> str:
        return json.dumps(asdict(self), default=str, indent=2)

    @classmethod
    def from_json(cls, s: str) -> "Workbook":
        d = json.loads(s)
        sheets = []
        for sd in d["sheets"]:
            sheets.append(Sheet(
                name=sd["name"], index=sd["index"],
                max_row=sd["max_row"], max_col=sd["max_col"],
                cells=[Cell(c["row"], c["col"], c["value"], c["display"],
                            Style(**c["style"]), c.get("hyperlink"), c.get("note"))
                       for c in sd["cells"]],
                merged=[MergedRange(**m) for m in sd["merged"]],
                col_widths={int(k): v for k, v in sd["col_widths"].items()},
                images=[ImageAsset(**i) for i in sd["images"]],
                screenshots=list(sd["screenshots"]),
            ))
        return cls(
            source=d["source"], sheets=sheets,
            media=[ImageAsset(**i) for i in d["media"]],
            has_charts=d.get("has_charts", False),
            has_embeddings=d.get("has_embeddings", False),
        )
```

Note: `col_widths` keys serialize to strings in JSON; `from_json` casts them back to int. `value` may be a non-JSON type (datetime) — `default=str` stringifies it on dump, and `display` carries the human form, so round-trip equality holds for the test's JSON-native values.

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_model.py -v`
Expected: PASS

- [ ] **Step 5: Write `tests/conftest.py` (fixture builders reused later)**

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import pytest


@pytest.fixture
def build_simple_xlsx(tmp_path):
    def _build(name="simple.xlsx"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Spec"
        ws["A1"] = "Parameter"; ws["B1"] = "280Ah"; ws["C1"] = "314Ah"
        ws["A1"].font = Font(bold=True)
        ws["A2"] = "Loading Level"; ws["B2"] = 25.40; ws["C2"] = 38.20
        ws["B2"].number_format = "0.00"; ws["C2"].number_format = "0.00"
        ws["A3"] = "N/P ratio"; ws["B3"] = 1.087; ws["C3"] = 1.086
        ws["B3"].number_format = "0.000"; ws["C3"].number_format = "0.000"
        ws.column_dimensions["A"].width = 18
        p = tmp_path / name
        wb.save(p)
        return str(p)
    return _build


@pytest.fixture
def build_rich_xlsx(tmp_path):
    """Merged header, a fill, a percent, a blank cell, and an embedded image."""
    def _build(name="rich.xlsx", with_image=True):
        from PIL import Image
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Cell Design"
        ws.merge_cells("A1:C1")
        ws["A1"] = "ESS LFP Cell Design"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A1"].fill = PatternFill("solid", fgColor="FFD966")
        ws["A2"] = "Capacity"; ws["B2"] = "280Ah"; ws["C2"] = "314Ah"
        ws["A3"] = "Margin"; ws["B3"] = 0.012; ws["B3"].number_format = "0.0%"
        ws["A4"] = "Note"; ws["C4"] = None  # intentional blank
        if with_image:
            img_path = tmp_path / "_e.png"
            Image.new("RGB", (40, 20), (10, 80, 160)).save(img_path)
            from openpyxl.drawing.image import Image as XLImage
            ws.add_image(XLImage(str(img_path)), "E2")
        p = tmp_path / name
        wb.save(p)
        return str(p)
    return _build
```

- [ ] **Step 6: Run full suite + commit**

Run: `pytest -q`
Expected: PASS (1 test).
```bash
git add -A && git commit -m "feat: data model with JSON round-trip + test fixtures"
```

---

### Task 3: `cleaners.format_cell_value`

**Files:**
- Create: `xl2word/xl2word/cleaners.py`, `xl2word/tests/test_cleaners.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `format_cell_value(value, number_format) -> str`, used by `extract_semantic` to fill `Cell.display`.

- [ ] **Step 1: Write the failing test** — `tests/test_cleaners.py`

```python
import datetime as dt
from xl2word.cleaners import format_cell_value

def test_none_is_empty():
    assert format_cell_value(None, "0.00") == ""

def test_general_passthrough():
    assert format_cell_value("Loading", None) == "Loading"
    assert format_cell_value("Loading", "General") == "Loading"

def test_decimals():
    assert format_cell_value(25.4, "0.00") == "25.40"
    assert format_cell_value(1.087, "0.000") == "1.087"

def test_integer():
    assert format_cell_value(142, "0") == "142"

def test_percent():
    assert format_cell_value(0.012, "0.0%") == "1.2%"
    assert format_cell_value(0.5, "0%") == "50%"

def test_thousands():
    assert format_cell_value(1190, "#,##0") == "1,190"

def test_date():
    assert format_cell_value(dt.datetime(2026, 4, 1), "yyyy-mm-dd") == "2026-04-01"

def test_fallback():
    assert format_cell_value(3.14159, "weird") == "3.14159"
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_cleaners.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write `cleaners.py`**

```python
from __future__ import annotations
import datetime as dt
import re


def format_cell_value(value, number_format: str | None) -> str:
    if value is None:
        return ""
    fmt = (number_format or "General").strip()
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%Y-%m-%d")
    if fmt in ("General", "@", ""):
        return str(value)
    if isinstance(value, (int, float)):
        try:
            if "%" in fmt:
                decimals = _decimals(fmt)
                return f"{value * 100:.{decimals}f}%"
            if "," in fmt and "." not in fmt:
                return f"{value:,.0f}"
            if "." in fmt:
                return f"{value:.{_decimals(fmt)}f}"
            if re.fullmatch(r"0+", fmt):
                return f"{int(round(value))}"
        except (ValueError, TypeError):
            return str(value)
    return str(value)


def _decimals(fmt: str) -> int:
    m = re.search(r"\.(0+)", fmt)
    return len(m.group(1)) if m else 0
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_cleaners.py -v`
Expected: PASS (8 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: number-format aware cell display formatting"
```

---

### Task 4: `extract_semantic` (openpyxl → Workbook)

**Files:**
- Create: `xl2word/xl2word/extract.py`
- Test: `xl2word/tests/test_extract.py`

**Interfaces:**
- Consumes: `model.*`, `cleaners.format_cell_value`, fixtures `build_simple_xlsx`/`build_rich_xlsx`.
- Produces: `extract_semantic(xlsx_path) -> Workbook` (cells, merges, styles, col widths; no media yet).

- [ ] **Step 1: Write the failing test** — `tests/test_extract.py`

```python
from xl2word.extract import extract_semantic

def test_semantic_captures_values_styles_merges(build_rich_xlsx):
    wb = extract_semantic(build_rich_xlsx())
    s = wb.sheets[0]
    assert s.name == "Cell Design"
    cells = {(c.row, c.col): c for c in s.cells}
    assert cells[(1, 1)].display == "ESS LFP Cell Design"
    assert cells[(1, 1)].style.bold is True
    assert cells[(1, 1)].style.fill == "FFD966"
    assert cells[(3, 2)].display == "1.2%"            # 0.012 with 0.0%
    assert any(m.min_row == 1 and m.max_col == 3 for m in s.merged)

def test_simple_col_width(build_simple_xlsx):
    wb = extract_semantic(build_simple_xlsx())
    assert wb.sheets[0].col_widths.get(1) == 18
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_extract.py -v`
Expected: FAIL — `ImportError: cannot import name 'extract_semantic'`

- [ ] **Step 3: Write `extract_semantic` in `extract.py`**

```python
from __future__ import annotations
import openpyxl
from openpyxl.utils import get_column_letter
from .model import Workbook, Sheet, Cell, Style, MergedRange
from .cleaners import format_cell_value


def _rgb(color) -> str | None:
    if color is None:
        return None
    rgb = getattr(color, "rgb", None)
    if not isinstance(rgb, str):
        return None
    if len(rgb) == 8:
        if rgb[:2] == "00":          # fully transparent / no-fill sentinel
            return None
        rgb = rgb[2:]
    return rgb.upper()


def _cell_style(c) -> Style:
    fill = None
    if c.fill is not None and c.fill.patternType:
        fill = _rgb(c.fill.fgColor)
    border = False
    if c.border is not None:
        border = any(getattr(c.border, side).style
                     for side in ("top", "bottom", "left", "right"))
    return Style(
        bold=bool(c.font and c.font.bold),
        italic=bool(c.font and c.font.italic),
        font_name=c.font.name if c.font else None,
        font_size=float(c.font.size) if c.font and c.font.size else None,
        font_color=_rgb(c.font.color) if c.font else None,
        fill=fill,
        align_h=c.alignment.horizontal if c.alignment else None,
        align_v=c.alignment.vertical if c.alignment else None,
        border=border,
        number_format=c.number_format,
    )


def extract_semantic(xlsx_path: str) -> Workbook:
    book = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheets: list[Sheet] = []
    for idx, ws in enumerate(book.worksheets):
        cells: list[Cell] = []
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and (c.fill is None or not c.fill.patternType):
                    continue  # skip truly empty/unstyled cells
                style = _cell_style(c)
                cells.append(Cell(
                    row=c.row, col=c.column, value=c.value,
                    display=format_cell_value(c.value, c.number_format),
                    style=style,
                    hyperlink=c.hyperlink.target if c.hyperlink else None,
                    note=c.comment.text if c.comment else None,
                ))
        merged = [MergedRange(r.min_row, r.min_col, r.max_row, r.max_col)
                  for r in ws.merged_cells.ranges]
        col_widths = {}
        for letter, dim in ws.column_dimensions.items():
            if dim.width:
                try:
                    col_widths[openpyxl.utils.column_index_from_string(letter)] = dim.width
                except ValueError:
                    pass
        sheets.append(Sheet(
            name=ws.title, index=idx,
            max_row=ws.max_row, max_col=ws.max_column,
            cells=cells, merged=merged, col_widths=col_widths,
        ))
    return Workbook(source=xlsx_path, sheets=sheets, media=[])
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_extract.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: semantic extraction of cells, styles, merges, widths"
```

---

### Task 5: `extract_media` (raw-zip sweep) + `extract_workbook` orchestrator

**Files:**
- Modify: `xl2word/xl2word/extract.py`
- Test: `xl2word/tests/test_extract.py`

**Interfaces:**
- Consumes: Task 4 output, `render.render_xlsx_to_images` (imported lazily; tolerated absent).
- Produces: `extract_media(xlsx_path, images_dir) -> list[ImageAsset]`; `extract_workbook(xlsx_path, out_dir, render=True) -> Workbook` that writes `out_dir/workbook.json`, `out_dir/images/`, `out_dir/screenshots/`.

- [ ] **Step 1: Write the failing test** — append to `tests/test_extract.py`

```python
import json, os
from xl2word.extract import extract_media, extract_workbook

def test_media_sweep_catches_embedded_image(build_rich_xlsx, tmp_path):
    imgs = extract_media(build_rich_xlsx(), str(tmp_path / "imgs"))
    assert len(imgs) >= 1
    assert os.path.exists(tmp_path / "imgs" / os.path.basename(imgs[0].path))

def test_extract_workbook_writes_json(build_simple_xlsx, tmp_path):
    out = str(tmp_path / "out")
    wb = extract_workbook(build_simple_xlsx(), out, render=False)
    assert os.path.exists(os.path.join(out, "workbook.json"))
    data = json.load(open(os.path.join(out, "workbook.json")))
    assert data["sheets"][0]["name"] == "Spec"
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_extract.py -v`
Expected: FAIL — `ImportError: cannot import name 'extract_media'`

- [ ] **Step 3: Add to `extract.py`**

```python
import os
import shutil
import zipfile
from .model import ImageAsset


def extract_media(xlsx_path: str, images_dir: str) -> list[ImageAsset]:
    os.makedirs(images_dir, exist_ok=True)
    assets: list[ImageAsset] = []
    with zipfile.ZipFile(xlsx_path) as z:
        for name in z.namelist():
            if name.startswith("xl/media/"):
                base = os.path.basename(name)
                dest = os.path.join(images_dir, base)
                with z.open(name) as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out)
                w = h = None
                try:
                    from PIL import Image
                    with Image.open(dest) as im:
                        w, h = im.size
                except Exception:
                    pass
                assets.append(ImageAsset(
                    id=base, path=os.path.join("images", base),
                    width_px=w, height_px=h, source="media",
                ))
    return assets


def _has_zip_dir(xlsx_path: str, prefix: str) -> bool:
    with zipfile.ZipFile(xlsx_path) as z:
        return any(n.startswith(prefix) for n in z.namelist())


def extract_workbook(xlsx_path: str, out_dir: str, render: bool = True) -> Workbook:
    os.makedirs(out_dir, exist_ok=True)
    images_dir = os.path.join(out_dir, "images")
    shots_dir = os.path.join(out_dir, "screenshots")
    wb = extract_semantic(xlsx_path)
    wb.media = extract_media(xlsx_path, images_dir)
    wb.has_charts = _has_zip_dir(xlsx_path, "xl/charts/")
    wb.has_embeddings = _has_zip_dir(xlsx_path, "xl/embeddings/")
    if render:
        try:
            from .render import render_xlsx_to_images
            shots = render_xlsx_to_images(xlsx_path, shots_dir)
            # Attach all page shots to sheet 0; multi-sheet refinement is a later concern.
            if wb.sheets and shots:
                wb.sheets[0].screenshots = [os.path.join("screenshots", os.path.basename(s))
                                            for s in shots]
        except Exception:
            pass  # rendering is best-effort; absence of soffice must not break extraction
    with open(os.path.join(out_dir, "workbook.json"), "w") as f:
        f.write(wb.to_json())
    return wb
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_extract.py -v`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: raw-zip media sweep + extract_workbook orchestrator"
```

---

### Task 6: `render.py` — xlsx/docx → PNG via soffice + PyMuPDF

**Files:**
- Create: `xl2word/xl2word/render.py`, `xl2word/tests/test_render.py`

**Interfaces:**
- Consumes: `soffice` on PATH (runtime), PyMuPDF.
- Produces: `render_xlsx_to_images`, `render_docx_to_images`, `RenderError`. Used by `extract_workbook` and `verify`.

- [ ] **Step 1: Write the failing test** — `tests/test_render.py`

```python
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
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_render.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xl2word.render'`

- [ ] **Step 3: Write `render.py`**

```python
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
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_render.py -v`
Expected: PASS (the xlsx test runs if soffice present, else SKIPPED; the error test always PASSES)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: LibreOffice + PyMuPDF rendering of xlsx/docx to PNG"
```

---

### Task 7: `fit.py` — column-width and orientation math

**Files:**
- Create: `xl2word/xl2word/fit.py`, `xl2word/tests/test_fit.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `EMU_PER_INCH`, `estimate_text_width_emu`, `natural_column_widths`, `fit_columns`, `choose_orientation`. Used by `docx_write`.

- [ ] **Step 1: Write the failing test** — `tests/test_fit.py`

```python
from xl2word.fit import (EMU_PER_INCH, natural_column_widths, fit_columns,
                         choose_orientation)

def test_natural_widths_track_content_length():
    w = natural_column_widths([["a", "wider content here"], ["b", "x"]], 10)
    assert w[1] > w[0]

def test_fit_scales_down_when_over_budget():
    natural = [5 * EMU_PER_INCH, 5 * EMU_PER_INCH]  # 10in natural
    fitted = fit_columns(natural, 6 * EMU_PER_INCH)  # 6in usable
    assert abs(sum(fitted) - 6 * EMU_PER_INCH) <= 2
    assert fitted[0] == fitted[1]

def test_fit_keeps_natural_when_under_budget():
    natural = [2 * EMU_PER_INCH, 2 * EMU_PER_INCH]
    assert fit_columns(natural, 6 * EMU_PER_INCH) == natural

def test_orientation_flips_for_wide():
    assert choose_orientation(9 * EMU_PER_INCH, 6 * EMU_PER_INCH, 9 * EMU_PER_INCH) == "landscape"
    assert choose_orientation(5 * EMU_PER_INCH, 6 * EMU_PER_INCH, 9 * EMU_PER_INCH) == "portrait"
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_fit.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write `fit.py`**

```python
from __future__ import annotations

EMU_PER_INCH = 914400
_MIN_COL_EMU = EMU_PER_INCH // 2  # 0.5in floor so columns never collapse


def estimate_text_width_emu(text: str, font_size_pt: float) -> int:
    # Average glyph advance ~0.5em; pad one em for cell margins.
    longest = max((len(line) for line in str(text).split("\n")), default=0)
    em_emu = font_size_pt * EMU_PER_INCH / 72.0
    return int((longest * 0.5 + 1.0) * em_emu)


def natural_column_widths(rows: list[list[str]], font_size_pt: float) -> list[int]:
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    widths = [_MIN_COL_EMU] * ncols
    for r in rows:
        for i, val in enumerate(r):
            widths[i] = max(widths[i], estimate_text_width_emu(val, font_size_pt))
    return widths


def fit_columns(natural: list[int], usable_width_emu: int) -> list[int]:
    total = sum(natural)
    if total <= usable_width_emu or total == 0:
        return list(natural)
    scale = usable_width_emu / total
    return [max(1, int(w * scale)) for w in natural]


def choose_orientation(natural_total_emu: int, portrait_usable_emu: int,
                       landscape_usable_emu: int) -> str:
    if natural_total_emu > portrait_usable_emu and \
       natural_total_emu <= landscape_usable_emu * 1.0:
        return "landscape"
    if natural_total_emu > portrait_usable_emu:
        return "landscape"
    return "portrait"
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_fit.py -v`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: column-fit and orientation math"
```

---

### Task 8: `layout.py` — Block / LayoutPlan + default_layout

**Files:**
- Create: `xl2word/xl2word/layout.py`, `xl2word/tests/test_layout.py`

**Interfaces:**
- Consumes: `model.Workbook`.
- Produces: `Block`, `LayoutPlan` (+ `to_json/from_json`), `default_layout(wb) -> LayoutPlan`. Consumed by `docx_write` and `cli`.

- [ ] **Step 1: Write the failing test** — `tests/test_layout.py`

```python
from xl2word.layout import Block, LayoutPlan, default_layout
from xl2word.model import Workbook, Sheet

def test_default_layout_section_per_sheet():
    wb = Workbook(source="x", sheets=[
        Sheet(name="Spec", index=0, max_row=3, max_col=3),
        Sheet(name="Mixing", index=1, max_row=2, max_col=2),
    ])
    plan = default_layout(wb)
    headings = [b.text for b in plan.blocks if b.kind == "heading"]
    assert headings == ["Spec", "Mixing"]
    assert any(b.kind == "table" and b.sheet == "Spec" for b in plan.blocks)

def test_layout_json_roundtrip():
    plan = LayoutPlan(title="T", blocks=[Block(kind="heading", text="H", level=1)])
    assert LayoutPlan.from_json(plan.to_json()) == plan
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_layout.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write `layout.py`**

```python
from __future__ import annotations
import json
from dataclasses import dataclass, field, asdict
from .model import Workbook


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


def default_layout(wb: Workbook) -> LayoutPlan:
    import os
    blocks: list[Block] = []
    for i, sheet in enumerate(wb.sheets):
        if i > 0:
            blocks.append(Block(kind="pagebreak"))
        blocks.append(Block(kind="heading", text=sheet.name, level=1))
        if sheet.cells:
            blocks.append(Block(kind="table", sheet=sheet.name))
        for img in sheet.images:
            blocks.append(Block(kind="image", path=img.path,
                                caption=os.path.basename(img.path)))
    title = os.path.splitext(os.path.basename(wb.source))[0]
    return LayoutPlan(title=title, blocks=blocks)
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_layout.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: LayoutPlan model and default layout"
```

---

### Task 9: `docx_write.write_docx` — editable, styled, fitted tables

**Files:**
- Create: `xl2word/xl2word/docx_write.py`, `xl2word/tests/test_docx_write.py`

**Interfaces:**
- Consumes: `model.Workbook`, `layout.LayoutPlan`, `fit.*`.
- Produces: `write_docx(wb, layout, out_path, images_dir) -> None`.

- [ ] **Step 1: Write the failing test** — `tests/test_docx_write.py`

```python
import os
from docx import Document
from xl2word.extract import extract_semantic
from xl2word.layout import default_layout
from xl2word.docx_write import write_docx

def test_write_docx_builds_editable_table(build_simple_xlsx, tmp_path):
    wb = extract_semantic(build_simple_xlsx())
    out = str(tmp_path / "out.docx")
    write_docx(wb, default_layout(wb), out, images_dir=str(tmp_path / "images"))
    assert os.path.exists(out)
    doc = Document(out)
    assert len(doc.tables) == 1
    t = doc.tables[0]
    flat = [c.text for row in t.rows for c in row.cells]
    assert "Loading Level" in flat
    assert "25.40" in flat            # display value preserved
    assert "Parameter" in flat

def test_merged_header_is_merged(build_rich_xlsx, tmp_path):
    wb = extract_semantic(build_rich_xlsx(with_image=False))
    out = str(tmp_path / "rich.docx")
    write_docx(wb, default_layout(wb), out, images_dir=str(tmp_path / "images"))
    doc = Document(out)
    row0 = doc.tables[0].rows[0]
    # A1:C1 merged -> the three grid cells resolve to one underlying cell
    assert row0.cells[0]._tc is row0.cells[2]._tc
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_docx_write.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xl2word.docx_write'`

- [ ] **Step 3: Write `docx_write.py`**

```python
from __future__ import annotations
import os
from docx import Document
from docx.shared import Emu, Pt, RGBColor
from docx.enum.section import WD_ORIENT
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
    rfonts = rpr.find(qn("w:rFonts")) or OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), _DEFAULT_FONT)
    if rpr.find(qn("w:rFonts")) is None:
        rpr.append(rfonts)
    return doc


def _set_orientation(section, orientation: str) -> None:
    if orientation == "landscape" and section.orientation != WD_ORIENT.LANDSCAPE:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width


def _usable_width_emu(section) -> int:
    return int(section.page_width - section.left_margin - section.right_margin)


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
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW")) or OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(int(width_emu / 635)))   # EMU -> twips
    tcW.set(qn("w:type"), "dxa")
    if tcPr.find(qn("w:tcW")) is None:
        tcPr.append(tcW)


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
    section = doc.sections[-1]
    text_rows = [[(cell.display if cell else "") for cell in row] for row in rows]
    natural = fit.natural_column_widths(text_rows, 10)
    _set_orientation(section, block.orientation)
    usable = _usable_width_emu(section)
    widths = fit.fit_columns(natural, usable)

    table = doc.add_table(rows=nrows, cols=ncols)
    table.style = "Table Grid"
    _fixed_layout(table)
    for gi, row in enumerate(rows):
        for gj, mcell in enumerate(row):
            _apply_cell(table.cell(gi, gj), mcell)
            _set_cell_width(table.cell(gi, gj), widths[gj])
    # Repeat header row across page breaks.
    _repeat_header(table.rows[0])
    # Apply merges that fall inside this region.
    for m in sheet.merged:
        if m.min_row >= r0 and m.max_row <= r1 and m.min_col >= c0 and m.max_col <= c1:
            a = table.cell(m.min_row - r0, m.min_col - c0)
            b = table.cell(m.max_row - r0, m.max_col - c0)
            a.merge(b)


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
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_docx_write.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: editable styled fitted Word tables with merges and images"
```

---

### Task 10: `verify.py` — render output + overflow detection

**Files:**
- Create: `xl2word/xl2word/verify.py`, `xl2word/tests/test_verify.py`

**Interfaces:**
- Consumes: `render.render_docx_to_images`, python-docx.
- Produces: `render_doc(docx_path, out_dir) -> list[str]`; `detect_overflow(docx_path) -> list[str]`.

- [ ] **Step 1: Write the failing test** — `tests/test_verify.py`

```python
from docx import Document
from docx.shared import Inches
from xl2word.verify import detect_overflow

def test_detect_overflow_flags_wide_table(tmp_path):
    doc = Document()
    section = doc.sections[0]
    usable_in = (section.page_width - section.left_margin - section.right_margin) / 914400
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).width = Inches(usable_in)         # each col = full usable width
    t.cell(0, 1).width = Inches(usable_in)         # total = 2x usable -> overflow
    p = str(tmp_path / "wide.docx")
    doc.save(p)
    issues = detect_overflow(p)
    assert any("wider than" in s.lower() for s in issues)

def test_no_overflow_when_narrow(tmp_path):
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).width = Inches(1)
    t.cell(0, 1).width = Inches(1)
    p = str(tmp_path / "narrow.docx")
    doc.save(p)
    assert detect_overflow(p) == []
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_verify.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write `verify.py`**

```python
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
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_verify.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: output rendering + deterministic table-overflow guard"
```

---

### Task 11: `cli.py` — deterministic end-to-end entry point

**Files:**
- Create: `xl2word/xl2word/cli.py`, `xl2word/tests/test_cli.py`

**Interfaces:**
- Consumes: `extract.extract_workbook`, `layout.default_layout`/`LayoutPlan`, `docx_write.write_docx`.
- Produces: `main(argv) -> int`; console script `xl2word`.

- [ ] **Step 1: Write the failing test** — `tests/test_cli.py`

```python
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
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_cli.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write `cli.py`**

```python
from __future__ import annotations
import argparse
import os
from .extract import extract_workbook
from .layout import default_layout, LayoutPlan
from .docx_write import write_docx


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="xl2word", description="Excel to Word converter.")
    ap.add_argument("input", help="path to .xlsx")
    ap.add_argument("-o", "--output", required=True, help="path to output .docx")
    ap.add_argument("--workdir", default=None, help="extraction working dir")
    ap.add_argument("--layout", default=None, help="layout.json to render against")
    ap.add_argument("--no-render", action="store_true", help="skip sheet screenshots")
    args = ap.parse_args(argv)

    workdir = args.workdir or (os.path.splitext(args.output)[0] + "_work")
    wb = extract_workbook(args.input, workdir, render=not args.no_render)
    if args.layout:
        plan = LayoutPlan.from_json(open(args.layout).read())
    else:
        plan = default_layout(wb)
    write_docx(wb, plan, args.output, images_dir=os.path.join(workdir, "images"))
    print(f"Wrote {args.output}")
    return 0
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_cli.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: CLI entry point wiring extract -> layout -> docx"
```

---

### Task 12: `SKILL.md` — the Claude orchestration skill

**Files:**
- Create: `xl2word/SKILL.md`

**Interfaces:**
- Consumes: the whole `xl2word` package via its CLI/modules.
- Produces: the user-invokable skill that runs preflight + the five stages.

- [ ] **Step 1: Write `SKILL.md`**

````markdown
---
name: xl2word
description: Use when converting an Excel/Google-Sheets .xlsx into a clean, publish-ready Word (.docx) — especially IBC recipe/spec sheets headed to customers. Produces editable, neatly-fitted tables, carries images, and verifies the result against a written layout contract.
---

# xl2word — Excel to Word

Convert any `.xlsx` into a clean, editable, neatly-laid-out `.docx`. Code captures the data; you (Claude) design the layout and verify the render.

## Preflight (do this first, every time)

1. **Model check.** This skill is built for maximum quality. Confirm you are running on **Opus 4.8 with the 1M context window at high effort**. If you are not, STOP and tell the user to relaunch, for example:
   `claude --model claude-opus-4-8[1m]` (and set high effort), ideally with `--dangerously-skip-permissions` so the multi-step pipeline runs without repeated prompts. The skip-permissions flag is a recommendation; the model is not.
2. **Tooling check.** Ensure `soffice` (LibreOffice) is on PATH and the package is installed (`pip install -e .` inside `xl2word/`). If `soffice` is missing, tell the user how to install it.

## The five stages

### 1. Extract (run the Python — do not eyeball the sheet)
Run: `python -m xl2word.cli "<input.xlsx>" -o "<out.docx>" --workdir "<work>"` to do a first deterministic pass, OR call extraction directly to inspect first:
`python -c "from xl2word.extract import extract_workbook; extract_workbook('<input.xlsx>', '<work>')"`
This writes `<work>/workbook.json`, `<work>/images/`, and `<work>/screenshots/`.

### 2. Understand
Read `<work>/workbook.json` AND open every image in `<work>/screenshots/`. Cross-reference them. Identify, per sheet: the real tables and their regions, group-header rows, titles/banners/footers, what each embedded image is, and any region the structured data did not capture cleanly (a chart, a dense graphic). Use the screenshot to resolve anything the cells alone are ambiguous about. If something is missing from the JSON but visible in the screenshot, write a small targeted extraction for just that piece.

### 3. Plan — write `layout.md` and `layout.json`
Write `<work>/layout.md` as the human-readable design contract, page by page, for example:
> Page 1: the Cell Specification comparison table (landscape, fitted within the page). Below it, the electrode-geometry image with caption.
> Page 2: the Mixing table (portrait).

Then write `<work>/layout.json` matching `LayoutPlan` (blocks: heading/table/image/pagebreak, with `sheet`, `region`, `orientation`, `path`, `caption`). For any region too visual to rebuild as an editable table, set the block to an `image` pointing at that region's screenshot — never drop content.

**Layout rule — one table, one page.** Aim to fit each table on a single page. Try whatever it takes to get there and still look clean: adjust column widths (wider or narrower), step the font down, switch the section to `landscape`, or split into logical sub-tables by region. If a table truly cannot be made neat within one page, letting it spill to a second page is acceptable, but one page is the default goal. It must always look clean and deliberate, never cramped or crappy.

**Writing rules for any text you add (titles, captions):** plain and human. No em dashes. None of: comprehensive, robust, leverage, delve, navigate, intricate, underscore, crucial, essential. The document carries the sheet's data — do not editorialize it.

### 4. Execute
Render against your plan:
`python -m xl2word.cli "<input.xlsx>" -o "<out.docx>" --workdir "<work>" --layout "<work>/layout.json"`

### 5. Verify against the contract (loop until clean)
Run `python -c "from xl2word.verify import render_doc, detect_overflow; print(detect_overflow('<out.docx>')); print(render_doc('<out.docx>', '<work>/verify'))"`.
First, fix anything `detect_overflow` reports. Then open every PNG in `<work>/verify/` and walk the whole document against `layout.md`: is each promised table present, does it fit the page width with no clipped columns, **does each table sit on a single page**, are merges intact, are images placed and uncut, and does every table look clean rather than cramped? For each table that spills onto a second page, try the one-page techniques in order — adjust column widths, step the font down, switch to landscape, or split by region — re-render, and re-check. For any other mismatch, adjust `layout.json` (orientation, region split, font, captions) and re-run stages 4-5. Repeat until the render matches `layout.md` and tables are one-page-and-neat wherever achievable. Do not stop at the first render.

## Reproducibility
Save the approved `layout.json` next to the source. For a new version of the same sheet, reuse it: re-run stage 1 then stage 4 with the saved `--layout`, then a quick stage 5 pass. Same structure, fresh data.
````

- [ ] **Step 2: Manual verification (no unit test for instructions)**

Confirm `SKILL.md` front matter parses (name + description present), every command references a real module path from this plan, and the writing rules match CLAUDE.md's AI-tells list.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: xl2word orchestration skill (preflight + five stages)"
```

---

### Task 13: README, example, end-to-end integration test

**Files:**
- Create: `xl2word/README.md`, `xl2word/tests/test_integration.py`, `xl2word/examples/` (generated artifacts ignored by git except a sample input)

**Interfaces:**
- Consumes: the full pipeline.
- Produces: a green end-to-end test and user-facing docs.

- [ ] **Step 1: Write the failing integration test** — `tests/test_integration.py`

```python
import os
from docx import Document
from xl2word.cli import main

def test_end_to_end_rich_sheet(build_rich_xlsx, tmp_path):
    src = build_rich_xlsx()                      # merged header + fill + percent + image
    out = str(tmp_path / "rich.docx")
    rc = main([src, "-o", out, "--workdir", str(tmp_path / "wd"), "--no-render"])
    assert rc == 0
    doc = Document(out)
    assert len(doc.tables) == 1
    flat = [c.text for row in doc.tables[0].rows for c in row.cells]
    assert "ESS LFP Cell Design" in flat
    assert "1.2%" in flat                        # percent display survived end-to-end
    # the embedded image was sweep-extracted and placed
    assert len(doc.inline_shapes) >= 1
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/test_integration.py -v`
Expected: FAIL — image not placed yet, because `default_layout` only emits image blocks from `sheet.images`, which `extract_semantic` leaves empty. This drives the final wiring fix.

- [ ] **Step 3: Wire swept media into the default layout**

In `xl2word/layout.py`, update `default_layout` to also emit image blocks from `wb.media` when a sheet has no per-sheet images:

```python
def default_layout(wb: Workbook) -> LayoutPlan:
    import os
    blocks: list[Block] = []
    for i, sheet in enumerate(wb.sheets):
        if i > 0:
            blocks.append(Block(kind="pagebreak"))
        blocks.append(Block(kind="heading", text=sheet.name, level=1))
        if sheet.cells:
            blocks.append(Block(kind="table", sheet=sheet.name))
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
```

Note: `extract_workbook` (not `extract_semantic`) populates `wb.media`; the integration test calls the CLI, which uses `extract_workbook`, so media is present. Update `tests/test_layout.py::test_default_layout_section_per_sheet` only if it now sees extra image blocks (it uses sheets with no media, so it stays green).

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/test_integration.py -v && pytest -q`
Expected: PASS (integration test + whole suite green)

- [ ] **Step 5: Write `README.md`**

```markdown
# xl2word

General Excel to Word converter. Deterministic Python captures everything in a
workbook (cells, styles, merges, embedded media, per-sheet screenshots); a Claude
skill designs the layout, renders editable and fitted Word tables, and verifies
the result against a written contract.

## Install
    pip install -e ".[dev]"
    # plus LibreOffice (provides `soffice`) for rendering

## CLI (deterministic path)
    xl2word input.xlsx -o output.docx
    xl2word input.xlsx -o output.docx --layout layout.json   # render a designed layout

## Skill (quality path)
Invoke the `xl2word` skill in Claude Code on Opus 4.8 1M, high effort. It runs the
five stages: extract, understand, write layout.md/layout.json, render, then
screenshot-verify against the contract.

## Tests
    pytest -q
```

- [ ] **Step 6: Add a sample input + commit**

```bash
mkdir -p examples
python -c "import tests.conftest" 2>/dev/null || true
git add -A
git commit -m "feat: end-to-end integration test, media-in-default-layout, README"
```

---

## Self-Review

**Spec coverage** (each SPEC.md section → task):
- Stage 1 universal extractor: semantic = Task 4, raw-zip media = Task 5, visual render = Task 6, wired in `extract_workbook` = Task 5. ✓
- Stage 2 understand, Stage 3 layout.md, Stage 5 verify-against-contract: encoded in `SKILL.md` Task 12 (Claude-driven, not unit code) + deterministic `detect_overflow` Task 10. ✓
- Stage 4 execute (editable, fitted, merges, images, CJK, headings): Task 9. ✓
- Reproducibility via saved layout.json: `cli --layout` Task 11 + SKILL.md Task 12. ✓
- Preflight (model/permissions/tooling): SKILL.md Task 12. ✓
- No AI tells: SKILL.md writing rules Task 12. ✓
- Repo structure, failure modes (overflow=Task 7/10, merges=Task 9, blanks=Task 3 `''`, CJK=Task 9 font, charts/exotic=Task 5 sweep + SKILL fallback): covered. ✓
- Open questions (real sheet, Google Sheets API, branding, PDF): left as open questions in SPEC; not implementation tasks. ✓

**Placeholder scan:** no TBD/TODO; every code/test step shows real code and a real command. ✓

**Type consistency:** `Workbook/Sheet/Cell/Style/MergedRange/ImageAsset` used identically across Tasks 2,4,5,8,9; `LayoutPlan/Block` identical across 8,9,11; `format_cell_value` signature stable 3→4; `render_*_to_images` stable 6→5,10; `fit.*` stable 7→9. `default_layout` signature unchanged in Task 13 (body extended only). ✓

**Known soft spots (acceptable for v1, caught by the Claude verify loop):** theme/indexed colors resolve to `None` (screenshot still shows true color); text-width estimate is heuristic (the verify loop catches any overflow `fit_columns` underestimates); multi-sheet screenshot attribution is approximate in `extract_workbook` (Claude re-derives per-sheet structure in Stage 2).
