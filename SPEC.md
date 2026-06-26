# xl2word — Excel → Word Conversion Skill

**Design spec** · 2026-06-26 · Ayush G (IBC engagement)
**Status:** approved design, pre-implementation

---

## Context

Kunal (CPO, IBC) keeps recipes and cell-design specs in Excel / Google Sheets as the working document, because editing values there is fast (bump a material spec from 7.7 to 8, add a property). When IBC publishes to customers and partners, they need a clean Word document instead. Today the conversion is the problem: a naive Excel-to-Word dump produces broken, misaligned tables that take heavy manual cleanup before anything is shippable.

This is a separate, general-purpose tool, not the one-time operator-log parser (`ai_parser_23-26_data/`, which owns the 200-sheet QC logs). Kunal will reuse this skill across different sheets whose contents we cannot predict in advance, so it must generalize to any spreadsheet rather than fit one shape. It lives in its own folder and ships to its own GitHub repo.

## Goal

Any `.xlsx` in, a clean and publish-ready `.docx` out, where every table is fitted neatly within the page, every image and element is preserved, and the text stays editable. Reproducible: drop in a new version of the same sheet and get the same clean document. Maximize the quality of the output; do not optimize for token cost.

## Non-goals (v1)

- The IBC operator logs (the parser already handles those).
- Recomputing live formulas, pivot tables, or pixel-faithful ActiveX / form controls. These are captured visually instead, never silently dropped.
- Per-customer bespoke templating beyond the saved layout contract.

## Principles

- **Code for capture, Claude for craft.** Deterministic Python guarantees the data and images come out complete and correct. Claude does the layout and design, where judgment is what separates a clean doc from a mangled one.
- **Completeness is a tested guarantee, not a per-run gamble.** The extractor is preset and content-agnostic, proven against diverse fixtures, so nothing is silently missed.
- **Two modalities beat one.** Claude gets both the exact structured data and a visual render of every sheet, and cross-references them.
- **The document carries the sheet's data, neatly laid out.** No AI-generated prose, no em dashes, no AI tells (per the CLAUDE.md list). Any unavoidable text (a title, a caption) is plain and human.

---

## Architecture — five stages

### Stage 1 — Extract (Python, universal)

A preset, content-agnostic extractor that reads the spreadsheet format itself rather than predicting contents, so it behaves identically on a recipe spec, a data dump, or a financial model. It reads three layers:

1. **Semantic (openpyxl):** every sheet; every used cell as both raw value and displayed (number-formatted) value; merged ranges; fills, fonts, borders, alignment; hyperlinks; cell notes; rich-text runs; column widths and row heights; conditional-format rules.
2. **Raw zip sweep:** the `.xlsx` is a zip, so the extractor walks `xl/media/`, `xl/drawings/`, `xl/charts/`, and `xl/embeddings/` directly. Every image, chart, and embedded object is captured even when openpyxl does not surface it through its API.
3. **Visual render:** LibreOffice headless (or HTML + Puppeteer) renders each sheet to an image, producing a pixel-true picture of every sheet.

**Output:** a structured JSON model (data + style + structure metadata), an images folder, and one or more per-sheet screenshots. Nothing is predicted; everything present is dumped.

The screenshot layer is the safety net that makes the extractor truly all-encompassing: anything that resists structured extraction is still captured visually. Between structured extraction and visual capture, nothing gets through both nets.

The extractor is tested against deliberately diverse fixtures (merged group headers, charts, CJK text, floating images, dense numeric data) to prove it generalizes rather than fitting one example.

### Stage 2 — Understand (Claude)

Claude reads all of Stage 1's output and makes sense of it before deciding anything: what kinds of data are present, which regions are tables, which rows are group headers, what the merges mean, what is a title / banner / footer, and where each image belongs. It cross-references the exact data against the screenshots, because the picture resolves structure that raw cells alone cannot.

### Stage 3 — Plan (`layout.md`, written by Claude)

Claude writes an explicit, page-by-page design contract describing exactly how the document should look, for example: "Page 1: the Cell Specification comparison table, landscape, fitted within the page margins; below it, the electrode-geometry image with its caption. Page 2: …". This `layout.md` is the ground truth the rest of the pipeline is measured against.

### Stage 4 — Execute (render to `.docx`)

Build the Word document against `layout.md`. Tables are native, editable Word tables (not pictures), with merges, fills, borders, and number formats preserved, and fitted to the page (column sizing, wrapping, landscape, or font step-down as needed). Images are placed inline with captions; hyperlinks are preserved; the confidential banner, header, and footer are carried so the doc is customer-ready. CJK and Latin text both render via a CJK-capable font.

**One table, one page (a rule, not a hard limit).** Each table should fit on a single page wherever it can be made to look clean. Claude tries whatever it takes to get there: adjusting column widths (wider or narrower), stepping the font down, switching that section to landscape, or tightening spacing. If a table genuinely cannot be made neat within one page, spilling to a second page is acceptable, but the default goal is one page. It must always look clean and deliberate, never cramped or crappy.

When a region is genuinely too visual to rebuild as an editable table (a chart, a dense graphic), Claude embeds that region's screenshot instead. Graceful fallback, nothing lost.

### Stage 5 — Verify against the plan (screenshot workflow, looped)

A final full pass: render the finished `.docx`, screenshot every page, and walk the whole document comparing it back to `layout.md`. Does page 1 hold that table, does it fit the page and sit on a single page, is the image present and uncut, are merges intact, and does every table look clean rather than cramped? Claude fixes any mismatch and re-renders, iterating until the document matches the contract. This makes verification objective (render vs a written spec) rather than a vibe check, and is what makes the output bulletproof.

---

## Reproducibility

On the first run for a given sheet, Claude designs the layout and the user approves it. The approved `layout.md` (section order, orientation, column mapping, titles, branding) is saved. New versions of the same sheet reflow fresh data into that same approved structure, so Kunal gets "drop new version in, same clean doc out" without per-run hand-tuning. The saved contract is also the one place to adjust the template once the structure is agreed.

## Preflight (model and permissions)

A skill cannot switch the running session's model or permission mode; those are set at launch. So the skill runs a preflight check at the start:

- If the current model is not **Opus 4.8 1M context** at **high effort**, it stops and prints the exact relaunch command, since output quality is the priority here. When the session already meets this (the common case), it proceeds.
- It recommends launching with `--dangerously-skip-permissions` so the multi-step pipeline (file writes, Python, screenshot tooling) runs without repeated prompts. This is a recommendation, not a hard gate.
- It verifies the screenshot toolchain (LibreOffice headless or Puppeteer) is available and sets it up if missing.

---

## Repo structure

```
IBC/xl2word/
  README.md
  pyproject.toml            # deps: openpyxl, python-docx, (Pillow), screenshot toolchain
  .gitignore
  SPEC.md                   # this file
  SKILL.md                  # the Claude skill: orchestrates the five stages + preflight
  xl2word/
    __init__.py
    extract.py              # Stage 1: semantic + raw-zip + visual render -> JSON + images + screenshots
    model.py                # the intermediate data model (dataclasses)
    render.py               # Stage 1c: xlsx -> per-sheet images (LibreOffice/Puppeteer)
    docx_write.py           # Stage 4: model + layout.md -> editable .docx
    fit.py                  # table fit-to-page logic (widths, wrap, orientation)
    verify.py               # Stage 5: render docx -> screenshots for comparison
    cli.py                  # `xl2word input.xlsx -o output.docx [--layout layout.md]`
  templates/
    default.docx            # base Word template: margins, fonts, heading + footer styles
  tests/
    test_extract.py
    test_fit.py
    test_docx_write.py
    fixtures/*.xlsx          # diverse fixtures proving generality
  examples/                  # sample input.xlsx -> output.docx
```

There are two entry paths. Running the **CLI alone** does Stage 1 (full extraction) and a best-effort Stage 4 render from a supplied `layout.md` — useful for scripting and reruns. The **skill** (`SKILL.md`) is the quality path: it adds the Claude understand, plan, and verify stages and the preflight, and is what guarantees the "every table fits, matches the contract" result. The Python modules are callable on their own so the skill orchestrates them rather than reimplementing them.

## Failure modes designed against

- **Wide-table overflow** (the main one): fit via column sizing, wrapping, landscape, or font step-down.
- **Merged cells:** reproduced as Word merges.
- **Sparse / blank cells:** render empty, never the string "None".
- **CJK text:** CJK-capable font so Hangul renders.
- **Packed cells** (a reading glued to its spec, ranges, multiple values): render the displayed value; keep raw available.
- **Charts and exotic objects:** captured by the raw-zip sweep and the screenshot net; embedded as images when not rebuildable.
- **Giant workbooks** (operator-log scale): out of scope, guarded against.

## Open questions (resolve during build / with Kunal)

1. **Real input sheet:** confirm mapping against an actual recipe/spec sheet from Kunal when available; until then, build against diverse fixtures including a reconstruction of the ESS LFP spec sheet.
2. **Google Sheets:** if the source is a live Google Sheet, `.xlsx` export is the default and loses almost nothing; direct Sheets API access would add cell notes and exact formatting if we want it.
3. **Branding assets:** source of the IBC/JASTECH logo, the CONFIDENTIAL banner, and the footer line, so the doc matches house style.
4. **Output format:** `.docx` is the deliverable; an optional PDF export for final publishing is a small add if wanted.
