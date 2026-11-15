# open-spreadsheet plugin

Skills and an MCP server that teach Claude Code and Antigravity to work in an [open-spreadsheet](../../README.md) workspace.

## What it adds

- **Skills**: `create-sheet` (new workbooks), `sheet-authoring` (the workbook format, formulas, number formats, layout), `current-sheet` (what the person has selected in the editor), `apply-notes` (carry out the notes people leave on cells).
- **MCP server** `open-spreadsheet-mcp` (stdio), 26 tools: `list_workbooks`, `read_range`, `list_assets`, `read_asset`, `write_note`, `render_chart`, `delete_asset`, `trace_cell`, `write_cells`, `set_style`, `set_layout`, `create_workbook`, `rename_workbook`, `duplicate_workbook`, `delete_workbook`, `add_sheet`, `rename_sheet`, `delete_sheet`, `insert_rows`, `insert_columns`, `list_notes`, `set_note`, `set_names`, `evaluate`, `check`, `get_current`. Writes recalculate, report the formula problems they leave, and refuse to overwrite a file that changed since it was read.
- **Rules**: the house rules for editing workbooks.

## Install

Claude Code:

```bash
# Claude Code
/plugin install open-spreadsheet --marketplace simonliu-ai-product/open-spreadsheet

# Google Antigravity (IDE, App, or the agy CLI)
agy plugin install https://github.com/simonliu-ai-product/open-spreadsheet
```

The MCP server runs with `npx -p @open-spreadsheet/core open-spreadsheet-mcp` and finds the workspace from the folder the agent works in (or `--cwd <dir>`).

> Until `@open-spreadsheet/core` is on npm, point the server at a local checkout instead: `node /path/to/open-spreadsheet/packages/core/mcp.js --cwd <workspace>`.
