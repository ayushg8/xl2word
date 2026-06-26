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
