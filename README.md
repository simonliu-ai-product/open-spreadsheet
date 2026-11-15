# open-spreadsheet

**Excel-like spreadsheets, built for coding agents.** Each workbook is one JSON file of what people typed — `"120"`, `"2026-01-05"`, `"=SUM(B2:B9)"` — so a person edits it in a real grid and an agent edits the file, and git diffs both.

[繁體中文](README.zh-TW.md)

## Highlights

### 🧮 Its own formula engine

No spreadsheet library underneath. The engine (`packages/engine`, zero dependencies) follows Excel's rules where compatibility bugs usually hide: type coercion, 15-significant-digit comparison, the 1900 date system with its leap-year bug, `-2^2 = 4`, ranges across sheets, whole columns (`A:A`), absolute references, and insert/delete of rows and columns that rewrite every formula (deleted cells become `#REF!`). Recalculation is a dependency graph with an explicit-stack topological order, so a 100,000-cell chain recalculates without overflowing; ranges are indexed by column, so a change looks only at the formulas that can see it. 337 functions so far — math and trigonometry, statistics (MAXIFS, PERCENTILE, RANK, CORREL, FORECAST), text (TEXTBEFORE/AFTER, TEXT, FIXED), dates (NETWORKDAYS, WORKDAY, DATEDIF, YEARFRAC), lookup (XLOOKUP, XMATCH, INDEX/MATCH), finance (PMT, RATE, IRR, XIRR, NPV), engineering (number bases, bitwise, complex numbers, CONVERT), probability distributions (NORM, T, CHISQ, F, BINOM, POISSON, GAMMA, BETA), the D-functions and the older names; see [docs/PLAN.md](docs/PLAN.md) for the full coverage plan against Microsoft's function catalogue.

### 🟩 A grid that behaves like Excel

- Formula bar and name box (type `B12` or `A1:C5` to jump), function autocomplete with signatures, coloured reference boxes while you type, click or arrow keys to point at references, `F4` to cycle `$`.
- Number formats (`#,##0`, `"NT$"#,##0`, `0.0%`, `yyyy-mm-dd`, custom codes with sections and `[Red]`); typing `2026/1/5`, `12%` or `$1,200` picks the format for you.
- Bold, italic, underline, strikethrough, text and fill colours, borders (all, outside, thick, single sides), alignment, wrap, merge cells, frozen panes, hidden rows and columns, column and row resize (double-click to fit).
- Right-click menus on cells, row and column headers and sheet tabs: cut/copy/paste (values only too), insert and delete rows and columns, clear, sort, merge, freeze.
- Sheet tabs: add, rename (double-click), duplicate, delete, reorder by dragging, tab colours.
- Defined names (`=SUM(Sales)`): define from the selection, a name manager, the name box jumps to a name, formula autocomplete offers names; they follow inserted rows and renamed sheets, and travel in .xlsx.
- Trace precedents and dependents, as Excel: dashed boxes and arrows show what a formula reads or which formulas read a cell, names resolved; a list names each one, other sheets included, and jumps to it. Esc or an edit removes the arrows.
- Conditional formats from Excel files: besides comparisons, text, duplicates, colour scales and data bars, a workbook's Top/Bottom N, above/below average and icon sets (arrows, traffic lights, symbols, flags, ratings) show in the grid and survive the round trip to .xlsx.
- Protected ranges: lock cells from the right-click menu or ⌘K; typing, pasting, clearing, filling, sorting and formatting into them are refused — and so are MCP writes, so an agent cannot change what a person locked.
- Checkboxes over TRUE/FALSE cells: insert from the right-click menu or ⌘K, click to tick, Space flips the selection; formulas such as `=COUNTIF(B2:B9, TRUE)` count them.
- Links: a typed web or mail address, or `=HYPERLINK(url, name)`, shows as a link; ⌘-click or *Open link* opens it (http, https and mailto only).
- Vertical alignment (top / middle / bottom) from the right-click menu or ⌘K, kept in .xlsx and print.
- Font size, as Excel's points (11 is the default): a toolbar menu and ⇧⌘> / ⇧⌘<; rows nobody resized grow to fit; kept in .xlsx and print.
- Custom sort: up to three columns, each ascending or descending, with or without a header row; ties keep their order.
- Show formulas (⌃`) puts every formula in place of its result; *Go to next error* walks the workbook's error cells and says what each error means (#DIV/0!, #REF!, #NAME?…).
- Data tools: remove duplicates, text to columns, paste transposed.
- AutoComplete from the entries already in a column, a format painter (⇧⌘C / ⇧⌘V or the brush), fill handle (series, dates, `Item 1 → Item 2`, formulas with shifted references), `⌘D` / `⌘R`, `⌘Enter` to fill a selection, copy and paste with Excel and Google Sheets (TSV), find and replace, undo and redo, zoom, and a status bar with sum, average, min, max and count.
- Charts — bar, line and pie — drawn from a range, floating over the grid: drag to move, pull the corner to resize, right-click to change type, title or range; they follow edits, filters and hidden rows.
- Conditional formatting (greater/less/between/equal, text contains, duplicates, colour scales, data bars), dropdown lists that refuse other input, and filters with a dropdown on each header — all stored in the workbook file.
- Notes on cells (right-click or ⇧F2): a corner mark, a hover popup and a list in the header — and a request your agent can carry out (`apply-notes`).
- ⌘K command palette; light and dark; English, 繁體中文, 简体中文, 日本語 and 한국어.

### 🖨️ Print or save as PDF

⌘P (or Download → Print) lays the sheet out as paged table: number formats, styles, merges and borders kept, hidden rows and columns left out, frozen header rows repeated on every page, charts after the table. The system print dialog saves it as a PDF.

### 📥 CSV and Excel in and out

Import `.csv`, `.tsv` or `.xlsx` (sheets, formulas — shared formulas included — number formats, fonts, fills, column widths, merges, frozen panes) and download `.xlsx` or CSV. The zip reader and writer are our own; nothing is uploaded anywhere.

### 📊 Feed open-dashboard

`open-spreadsheet export <workbook> [--sheet] [--range A1:F200] [--format csv|json] [--out file]` writes a sheet's computed values — formulas evaluated, numbers plain, dates as ISO dates, the first row as the header. An [open-dashboard](../open-dashboard) collector runs it and a `csv` source reads the file, so a dashboard charts what people keep in the grid:

```ts
// open-dashboard.config.ts
datasources: { budget: { type: 'csv', tables: { orders: { file: 'data/orders.csv' } } } },
collectors: {
  orders: { run: 'npx -p @open-spreadsheet/core open-spreadsheet export sales-report --range A1:F200 --out data/orders.csv --root ../my-sheets', source: 'budget', every: '5m' },
},
```

### 🤖 Agent-native

- `init` writes `AGENTS.md`, `CLAUDE.md`, `.mcp.json` and four skills (`create-sheet`, `sheet-authoring`, `current-sheet`, `apply-notes`); `sync-skills` refreshes them.
- The editor records what is selected in `node_modules/.open-spreadsheet/current.json`, so "this cell" means something to an agent.
- `open-spreadsheet check` lists formulas that show `#NAME?`, `#REF!`, `#CIRC!` or `#VALUE!`, and malformed cells, styles and merges.
- `open-spreadsheet-mcp` is an MCP server over stdio with 21 tools — `read_range`, `trace_cell`, `write_cells`, `set_style`, `set_layout`, `create_workbook`, `rename_workbook`, `duplicate_workbook`, `delete_workbook` (to `sheets/.trash/`, never erased), sheet and row/column operations, `list_notes`, `set_note`, `set_names`, `evaluate`, `check`, `get_current`. Every write recalculates, reports the problems it leaves, and is refused (SHA-256 check) if the file changed since it was read. The viewer saves the same way, so an agent and a person never overwrite each other — and an agent's edit shows up in the open grid at once.
- A plugin for Claude Code and Antigravity in [plugins/open-spreadsheet](plugins/open-spreadsheet):

```bash
# Claude Code
/plugin install open-spreadsheet --marketplace simonliu-ai-product/open-spreadsheet

# Google Antigravity
agy plugin install https://github.com/simonliu-ai-product/open-spreadsheet
```

## Get started

```bash
npx -p @open-spreadsheet/core open-spreadsheet init my-sheets
cd my-sheets && pnpm install && pnpm dev    # http://localhost:5373
```

Or open the folder in [CoStaff Desktop](../open-desktop).

## The file contract

```json
{
  "title": "2026 budget",
  "sheets": [
    {
      "name": "Budget",
      "cells": { "A1": "Item", "B1": "Amount", "A2": "Rent", "B2": "22000", "B5": "=SUM(B2:B4)" },
      "styles": { "A1": { "b": true, "bg": "#e8f0e9" }, "B5": { "b": true, "fmt": "#,##0" } },
      "cols": { "A": 160 },
      "merges": ["A7:C7"],
      "freeze": { "rows": 1, "cols": 0 }
    }
  ]
}
```

## Repo layout

| Path | What |
| --- | --- |
| `packages/engine` | `@open-spreadsheet/engine` — parser, evaluator, functions, number formats, dependency graph. |
| `packages/core` | `@open-spreadsheet/core` — the CLI (`dev`, `init`, `check`, `sync-skills`), the dev server, the grid editor (`app/`), the MCP server (`mcp.js`) and the skills. |
| `plugins/open-spreadsheet` | The agent plugin for Claude Code and Antigravity. |
| `docs/PLAN.md` | Architecture, function coverage, milestones. |

## Development

```bash
pnpm test        # engine, ops and MCP tests (vitest)
pnpm typecheck
pnpm bench       # recalculation benchmarks
node scripts/ui-e2e.mjs <out-dir>   # drive the grid in Chromium, with screenshots
```

## License

MIT
