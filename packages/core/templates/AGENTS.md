# open-spreadsheet workspace

Excel-like spreadsheets kept as plain JSON, so people edit them in the grid and agents edit the files.

## Commands

```bash
pnpm dev                       # open-spreadsheet dev: the grid editor at http://localhost:5373
pnpm check                     # open-spreadsheet check: broken formulas, bad references, malformed files
npx open-spreadsheet sync-skills   # refresh the skills in .claude/skills and .agents/skills
npx open-spreadsheet-mcp       # the MCP server over stdio (also in .mcp.json)
npx open-spreadsheet export <id> --range A1:F200 --out data.csv   # computed values for open-dashboard and other tools
```

## Files

- `sheets/<id>/workbook.json` — one workbook. `<id>` is lowercase letters, numbers and dashes.
- `open-spreadsheet.config.json` — marks the folder as a workspace.
- `node_modules/.open-spreadsheet/current.json` — what the person has open and selected right now (see the `current-sheet` skill).

## Workbook format

```json
{
  "title": "2026 budget",
  "sheets": [
    {
      "name": "Budget",
      "cells": { "A1": "Item", "B1": "Amount", "A2": "Rent", "B2": "22000", "B5": "=SUM(B2:B4)" },
      "styles": { "A1": { "b": true, "bg": "#e8f0e9" }, "B5": { "b": true, "fmt": "#,##0" } },
      "cols": { "A": 160 },
      "rows": { "1": 30 },
      "merges": ["A7:C7"],
      "freeze": { "rows": 1, "cols": 0 },
      "color": "#43a047",
      "notes": { "B5": "Check this against the bank statement" }
    }
  ]
}
```

- **Cells hold what a person would type, always as strings**: `"120"`, `"2026-01-05"`, `"12%"`, `"=SUM(B2:B4)"`. A leading apostrophe keeps text that looks like a number: `"'00123"`.
- **Formulas** follow Excel: A1 references, `$` for absolute, `Sheet2!A1` or `'My sheet'!A1` across sheets, the usual operators and functions (SUM, SUMIF(S), COUNTIF(S), AVERAGE, ROUND, IF/IFS/IFERROR, VLOOKUP/XLOOKUP/INDEX/MATCH, TEXT, DATE/TODAY/EDATE/EOMONTH, PMT/NPV and more). An unknown function shows `#NAME?`.
- **Styles** per cell: `b`, `i`, `u`, `s` (strike), `align` (`left`/`center`/`right`), `wrap`, `valign` (`top`/`middle`/`bottom`), `check` (a checkbox over a `TRUE`/`FALSE` value), `size` (font size in points, 11 by default), `color`, `bg`, and `fmt`, an Excel number format such as `#,##0`, `0.0%`, `"NT$"#,##0`, `yyyy-mm-dd`.
- Sizes are pixels: `cols` by letter, `rows` by 1-based row number.
- **Defined names** sit at the top level: `"names": { "Sales": "Orders!$F$2:$F$200" }`, then `=SUM(Sales)`.
- **Conditional formats, dropdown lists and filters** live beside the cells too (`conditional`, `validation`, `filter`); see the `sheet-authoring` skill.
- **Assets** sit beside each workbook in `sheets/<id>/assets/` (and shared in `assets/`): chart images and Markdown notes. **Read a workbook's notes and the shared ones before you change it** — they say what the numbers mean and what to keep. Charts are never on the grid: make one with the MCP `render_chart` tool (an SVG in Assets). Keep what you learn with `write_note`.
- **Protected ranges** (`"protected": ["A1:D1"]` on a sheet) are cells a person locked. Do not change them, by MCP or by editing the JSON; if a task needs them changed, ask the person to unprotect them in the editor.
- **Notes** are comments people leave on cells, often requests for you. "Apply the notes" means: do what each asks, then remove it (the `apply-notes` skill).

## Working rules

1. Prefer the MCP tools (`read_range`, `write_cells`, `set_style`, `check`) when they are available; they recalculate and refuse to overwrite a file someone changed meanwhile.
2. Editing the JSON directly is fine too: keep it valid JSON, keep every input a string, and keep the sheet order.
3. Put totals and derived numbers in formulas, not typed results, so they stay right when inputs change.
4. After any change, run `open-spreadsheet check` (or the `check` tool) and fix what it reports.
5. "This cell", "the selected range", "this sheet": read `node_modules/.open-spreadsheet/current.json` first (it includes the cell's note).
6. Before changing or deleting a cell, `trace_cell` shows which formulas read it; "where does this number come from" is the same tool's `precedents`.
7. To try a variation, `duplicate_workbook` first and edit the copy. `delete_workbook` moves a workbook to `sheets/.trash/`; use it only when asked.
