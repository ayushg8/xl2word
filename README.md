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
