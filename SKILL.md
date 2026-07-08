---
name: xl2word
description: Use when converting an Excel/Google-Sheets .xlsx into clean, customer-facing Word (.docx) process documentation. Claude reads and understands each sheet, then composes it like a person writing a spec document (settings as key/value lists, note-lists as bullets, tabular data as fitted tables) rather than scraping every block into a grid. Editable tables, one table per page, compact.
---

# xl2word: Excel to Word

Turn a spreadsheet into a **customer-facing process-spec document**. The Excel stays the team's working/editing tool; the Word doc is what they show a customer to document their process. So the goal is not a faithful dump of the sheet, it is a **document a person would be proud to hand over**: dense, organized, and readable.

The division of labor is the whole point:
- **Python captures** every value/style/merge (reliable, deterministic).
- **You (Claude) understand** each sheet and decide how to present it (taste, structure).
- **Python renders** your plan into the .docx (deterministic).

Do not try to make Python guess structure. You read the sheet and compose it.

## Preflight (every time)

1. **Model check.** Built for maximum quality; confirm you are on **Opus 4.8, 1M context, high effort**. If not, stop and tell the user to relaunch (`claude --model claude-opus-4-8[1m]`, high effort, ideally `--dangerously-skip-permissions`).
2. **Tooling.** `soffice` (LibreOffice) on PATH and the package installed (`pip install -e .` inside the repo). If `soffice` is missing, tell the user how to install it.

## 1. Extract (Python captures the data)

`python -c "from xl2word.extract import extract_workbook; extract_workbook('<input.xlsx>', '<work>')"`
writes `<work>/workbook.json` (every cell value, style, merge) plus `images/` and `screenshots/`.

## 2. Understand + compose each sheet (this is the job)

For **each sheet**, read its cells (values, which are bold, which are filled, the merged ranges) and work out what each region actually *is*, then emit an ordered list of blocks describing how to present it. On a large workbook, understand sheets in parallel (one subagent per sheet), then merge the block lists.

Compose with this vocabulary. Pick the block that matches what the content *is*:

- **heading** (`level` 2): a section title. Turn bold section banners ("Pre-set Conditions", "Slitting Results") into level-2 headings. Do **not** emit the sheet's own name or its big ALL-CAPS title banner; the level-1 sheet heading is added for you.
- **prose** (`text`): one short plain-English sentence introducing a sheet or section. Use sparingly, only where it adds clarity.
- **keyvalue** (`pairs: [[key, value], ...]`): a small set of single settings/parameters (roughly a 2-column, ≤ ~12-row parameter→value block). Reads as a clean settings list instead of a grid. This is the right choice for most "Pre-set Conditions" / process-condition blocks.
- **bullets** (`items: [...]`): a list of notes, instructions, or maintenance steps that are really prose, not a table. **This is the biggest readability win.** Merge a sentence split across several rows into one bullet and lightly tidy connective wording.
- **note** (`text`): one short shaded callout (a batch-totals summary line, a single important warning).
- **table** (`region: [r0, c0, r1, c1]`): genuinely tabular data with multiple value columns (a step recipe, a material list, a metrics table with target/range, a results grid). Reference the source rectangle including its header row(s); the renderer pulls the exact cells, so numbers cannot be altered. Set `orientation: "landscape"` only for a genuinely wide table, else `"portrait"`.

**Worked example — a mixing sheet:** room conditions → `keyvalue`; the ingredient list (Category/Product/mass/%) → `table`; a "batch totals" line → `note`; the Measurement/Target/Range block → `table`; the six "Additional notes" sentences → `bullets` (one per note, split-row sentences merged); the 14-step Equipment/Step/Process/RPM/Time recipe → `table`; Cleaning and Maintenance sentences → `bullets`.

### Hard rules

- **Never alter data.** Every number, unit, tolerance (±), spec, and material/supplier name stays exact. For anything dense or numeric, use a `table` region (zero transcription risk). Only transcribe into `keyvalue`/`bullets`/`note`/`prose`, and there you may merge split cells and tidy grammar but never invent, drop, round, or change a value. Do not add facts that are not in the sheet.
- **Writing style for any text you author** (prose, bullets, notes, captions, headings): plain and human. No em dashes. None of: comprehensive, robust, leverage, delve, navigate, intricate, underscore, crucial, essential. (The renderer also strips em dashes and you can check with `xl2word.cleaners.find_ai_tells`, but write it right in the first place.)
- **Leave out what is not customer-facing.** Skip empty sheets, `(WIP)` / in-progress tabs, and internal trackers (action-item logs, Q&A). Keep the document to the process spec.
- **Compact, no ceremony.** No cover page and no multi-page table of contents. Every table fits on one page (the renderer shrinks the font to fit width and height and keeps each table whole). Aim for a short document.

## 3. Render

Write `<work>/layout.json` as a single `LayoutPlan`: `{"title": "...", "blocks": [...]}`. For each kept sheet, emit `{"kind":"heading","text":"<Sheet Name>","level":1}` then that sheet's composed blocks. **Every `table` block must carry `"sheet": "<Sheet Name>"`** as well as its `region`, or the renderer skips it. Example blocks:

```json
{"kind":"heading","text":"Anode Mix","level":1}
{"kind":"prose","text":"Anode slurry is mixed in a 100 L single-vessel planetary mixer."}
{"kind":"heading","text":"Pre-set Conditions","level":2}
{"kind":"keyvalue","pairs":[["Room Dew Point (°C)","< -20"],["Room Temperature (°C)","23"]]}
{"kind":"heading","text":"Raw Materials","level":2}
{"kind":"table","sheet":"Anode Mix","region":[6,7,12,14],"orientation":"portrait"}
{"kind":"note","text":"Batch totals: 50.0% solid content, 83.30 kg total mass, 59.5 L total volume."}
{"kind":"bullets","items":["Viscosity is the primary metric to achieve within target value.","Keep the chiller connected to the main tank and running."]}
```

Then render:

`python -c "from xl2word.extract import extract_workbook; from xl2word.layout import LayoutPlan; from xl2word.docx_write import write_docx; wb=extract_workbook('<input.xlsx>','<work>',render=False); write_docx(wb, LayoutPlan.from_json(open('<work>/layout.json').read()), '<out.docx>', images_dir='<work>/images')"`

## 4. Verify (loop until clean)

1. **Data integrity.** For every number you authored into a `keyvalue`/`bullets`/`note`, confirm it appears in the source sheet's cells. A number in the doc that is not in the sheet is a bug; fix it.
2. **Visual.** Render to images: `python -c "from xl2word.verify import render_doc; print(render_doc('<out.docx>', '<work>/verify'))"`. Walk every page: is each section present, does each table sit on one page with no clipped columns, are settings/notes rendered as key-value/bullets (not crammed grids), are merges intact, and does it read like documentation rather than a scrape? Fix the plan and re-render. Do not stop at the first pass.

## Reproducibility

Save the approved `layout.json` next to the source. For a new revision of the same workbook, reuse it: re-extract, then render with the saved plan and do a quick verify. Same structure, fresh data.
