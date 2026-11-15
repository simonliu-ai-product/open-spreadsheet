---
name: create-sheet
description: Use this skill when the user wants a new spreadsheet in this open-spreadsheet workspace — a budget, tracker, schedule, invoice, report table, or "put this data in a sheet". Triggers on "make a spreadsheet", "track X in a sheet", "build a table of", "turn this CSV into a sheet". Do NOT use for editing the framework itself.
---

# Creating a workbook

1. **Name it.** Pick an id of lowercase letters, numbers and dashes (`q1-sales`), and a human title (`Q1 sales`).
2. **Plan the sheets.** One sheet per kind of record (Orders, Products), plus a Summary sheet when there are totals across them.
3. **Write `sheets/<id>/workbook.json`** (or call the MCP `create_workbook` then `write_cells`):
   - Row 1 is the header row, bold with a light fill, frozen (`"freeze": { "rows": 1, "cols": 0 }`).
   - Inputs go in as typed strings; dates as `yyyy-mm-dd`.
   - Totals, shares, lookups and status columns are **formulas** (`=SUM(...)`, `=C2/$C$10`, `=IF(...)`), never typed results.
   - Format numbers with `fmt` (`#,##0`, `"NT$"#,##0`, `0.0%`, `yyyy-mm-dd`).
   - Widen label columns (`"cols": { "A": 140 }`).
4. **Check it.** Run `open-spreadsheet check` (or the `check` tool) and fix every error it lists.
5. **Tell the person** the workbook's title and where to see it (`/#/w/<id>` in the editor).

See the `sheet-authoring` skill for the full format and the list of functions.
