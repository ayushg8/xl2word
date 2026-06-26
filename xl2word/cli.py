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
        with open(args.layout) as f:
            plan = LayoutPlan.from_json(f.read())
    else:
        plan = default_layout(wb)
    write_docx(wb, plan, args.output, images_dir=os.path.join(workdir, "images"))
    print(f"Wrote {args.output}")
    return 0
