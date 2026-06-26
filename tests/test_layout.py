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
