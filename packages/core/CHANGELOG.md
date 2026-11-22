# @open-spreadsheet/core

## 0.5.0

### Minor Changes

- [#5](https://github.com/simonliu-ai-product/open-spreadsheet/pull/5) [`9ec57b1`](https://github.com/simonliu-ai-product/open-spreadsheet/commit/9ec57b1faf03005010be4f3da765382e5f184a37) Thanks [@LiuYuWei](https://github.com/LiuYuWei)! - An agent beside the grid: the editor has a Codex panel (⌘J) that works in the workspace on your own Codex sign-in. Each turn knows the workbook, sheet and selection, changes workbooks through open-spreadsheet's own tools so formulas and references stay consistent, and when it ends the workbooks it changed are checked and the findings offered back. Connect installs `@openai/codex-sdk` into the workspace when you ask. Conversations stay in `node_modules/.open-spreadsheet/`; `"agent": false` in the config turns the panel off, and a host app that declares the open-x `agent` capability hides it.
