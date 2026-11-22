---
"@open-spreadsheet/core": minor
---

An agent beside the grid: the editor has a Codex panel (⌘J) that works in the workspace on your own Codex sign-in. Each turn knows the workbook, sheet and selection, changes workbooks through open-spreadsheet's own tools so formulas and references stay consistent, and when it ends the workbooks it changed are checked and the findings offered back. Connect installs `@openai/codex-sdk` into the workspace when you ask. Conversations stay in `node_modules/.open-spreadsheet/`; `"agent": false` in the config turns the panel off, and a host app that declares the open-x `agent` capability hides it.
