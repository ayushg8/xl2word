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
        if sheet.cells or sheet.max_row > 0:
            blocks.append(Block(kind="table", sheet=sheet.name))
        for img in sheet.images:
            blocks.append(Block(kind="image", path=img.path,
                                caption=os.path.basename(img.path)))
    title = os.path.splitext(os.path.basename(wb.source))[0]
    return LayoutPlan(title=title, blocks=blocks)
