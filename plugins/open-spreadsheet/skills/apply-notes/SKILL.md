---
name: apply-notes
description: Use this skill when the user asks you to apply, process, handle or clear the notes they left on cells in the open-spreadsheet editor — "apply the notes", "do what my notes say", "I left comments on the sheet", "/apply-notes". It finds every note, makes the change each one asks for, then removes that note. Do NOT use it to write new notes or for general editing.
---

# Applying notes

People leave notes on cells in the editor (right-click → Add note, or ⇧F2). A note is stored next to the cells in `sheets/<id>/workbook.json`:

```json
{ "name": "Budget", "cells": { "F4": "=SUM(C4:E4)" }, "notes": { "F4": "Show this as a share of the total, one decimal" } }
```

1. **Find them.** With the MCP server, call `list_notes` (all workbooks, or one). Otherwise read the `notes` of every sheet in `sheets/*/workbook.json`.
2. **Sort them.** A note that asks for something ("make this a percentage", "add a total under this column", "this should use the March numbers") is a request. A note that only explains ("source: 2025 annual report") is documentation: leave it.
3. **Do each request** on the cell the note is on and whatever it names — cells, formulas, styles, formats, layout — following the `sheet-authoring` skill. Keep totals as formulas.
4. **Remove each note you carried out** (`set_note` with empty `text`, or delete the key from `notes`). If a request is unclear or impossible, leave the note and say why.
5. **Check.** Run `open-spreadsheet check` (or the `check` tool) and fix what it reports.
6. **Report** which notes you applied, which you left and why.

The editor reloads as you write, so the person sees each change and the note disappear.
