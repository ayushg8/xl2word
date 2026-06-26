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
        if rgb == "00000000":        # no-color sentinel
            return None
        rgb = rgb[2:]                # strip ARGB alpha prefix
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
