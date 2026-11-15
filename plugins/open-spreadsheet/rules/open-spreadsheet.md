# open-spreadsheet

In a folder with `open-spreadsheet.config.json`, workbooks live in `sheets/<id>/workbook.json`.

- Cells hold what a person would type, always as strings; formulas start with `=` and follow Excel.
- Put totals and derived numbers in formulas, never typed results.
- Prefer the open-spreadsheet MCP tools (`read_range`, `write_cells`, `set_style`, `set_layout`, `check`); pass the `hash` from a read as `expected_hash` so you never overwrite the person's edit.
- "This cell" or "the selection" means what `get_current` (or `node_modules/.open-spreadsheet/current.json`) says.
- Finish with `check` and fix every error it reports.
