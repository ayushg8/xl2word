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
