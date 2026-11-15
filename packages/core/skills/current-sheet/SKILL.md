---
name: current-sheet
description: Resolve which workbook, sheet, cell or range the person is looking at in the open-spreadsheet editor. Consult this whenever the user says "this cell", "the selected range", "this sheet", "this table", "here", or otherwise points at something on screen without naming it. Re-read the file at the start of every such turn — the person moves between turns, so an earlier reading is stale.
---

# What is on screen

The editor writes `node_modules/.open-spreadsheet/current.json` every time the selection moves:

```json
{
  "workbook": "budget-2026",
  "title": "2026 budget",
  "sheet": "Budget",
  "active": "C4",
  "selection": "B2:D9",
  "input": "=SUM(C2:C3)",
  "value": 300,
  "note": "Should include the March refund",
  "file": "sheets/budget-2026/workbook.json",
  "updatedAt": "2026-11-15T08:30:00.000Z"
}
```

- `active` is the cell with the cursor; `selection` is the highlighted range (equal to `active` for one cell).
- `input` is what the cell holds (a formula starts with `=`); `value` is what it computes to; `note` is the note on it, if any.
- Edit `file` (or use the MCP tools with `workbook` and `sheet`) to change what the person sees; the editor reloads at once.

With the MCP server, the `get_current` tool returns the same object.

If the file is missing, the editor is not open: ask which workbook and cells they mean.
