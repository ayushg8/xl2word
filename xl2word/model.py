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
