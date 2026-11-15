# open-spreadsheet

類 Excel 的試算表框架（open-x 系列）。計算引擎和格子介面都自己寫，不用 Univer、HyperFormula 等套件。規劃、函數支援評估和里程碑見 `docs/PLAN.md`。

## 常用指令

```bash
pnpm test        # 引擎測試（vitest）
pnpm typecheck   # engine + core/app
pnpm bench       # 重算效能
node scripts/ui-e2e.mjs <輸出資料夾>   # Chromium 端對端檢查（99 項）＋截圖
# 在工作區內（例如 ../playground/open-spreadsheet-sample）
pnpm dev         # open-spreadsheet dev，預設 port 5373
```

## 結構

- `packages/engine/src/`：零依賴計算引擎。`values.ts`（值模型、轉型、比較）、`parse.ts`（剖析器）、`serialize.ts`（序列化、列欄位移）、`functions.ts`（函數表）、`dates.ts`、`address.ts`、`workbook.ts`（相依圖與重算）
- `packages/core/`：`bin.js`（CLI：dev／init／check／sync-skills）、`server.js`（Vite dev server、`/__api/workbooks`、`current.json`）、`ops.js`（列出、讀寫含 SHA-256 樂觀鎖、check，dev server／CLI／MCP 共用）、`mcp.js`（stdio MCP server，手寫 JSON-RPC、無依賴）、`skills/` 與 `templates/AGENTS.md`（init 寫進工作區）、`app/`（格子介面，直接 import 引擎原始碼）
  - `app/book.ts` 編輯器（格子、凍結窗格、合併、格式、選單、填滿、尋找、歷史、分頁）、`home.ts` 首頁、`io.ts`（CSV／xlsx 與 zip）、`ui.ts`（選單、對話框、⌘K、主題）、`i18n.ts`＋`i18n-more.ts`（五種語言，新字串五種都要補）、`templates.ts`、`help.ts`（函數說明）
- `packages/open-spreadsheet/`：npm 上沒有 scope 的 `open-spreadsheet`，只是 `@open-spreadsheet/core` CLI 的短名稱（`npx open-spreadsheet init`），也避免這個名稱被別人搶註
- `plugins/open-spreadsheet/`：Claude Code／Antigravity plugin；`skills/` 是 `packages/core/skills/` 的複本，`agent.test.ts` 檢查兩邊一致
- 工作區格式：`open-spreadsheet.config.json` 加 `sheets/<id>/workbook.json`，存使用者輸入的原文

## 慣例

- 引擎行為以 Excel 為準：先查微軟文件，有疑問就寫測試記下 Excel 的結果。
- 新增函數時加進 `functions.ts` 的函數表，並在 `engine.test.ts` 補測試。
- 引擎用 Node 的 strip-types 執行，不能用 TypeScript 的 parameter properties、enum 等需要轉譯的語法。
- core 一律用套件名 `@open-spreadsheet/engine` 引用引擎，不要寫 `../engine/src` 相對路徑。Node 不會替 `node_modules` 裡的 `.ts` 去型別，所以發佈時引擎走 `publishConfig.exports` 指向 `dist/`；`dist/` 由 `pnpm pack` 前的 `scripts/build.mjs` 產生。repo 內仍直接跑 `src/`。`agent.test.ts` 的 packaging 測試會檢查這兩件事。
- 重算不可遞迴（長鏈會 stack overflow），維持明確堆疊。
- 編輯器裡 `active` 是作用中儲存格（不動），`anchor` 是選取範圍會移動的那一角。
- 所有寫檔都走 `ops.js` 的 `writeWorkbook`（佇列＋SHA-256）；檢視器存檔帶 If-Match，412 時重新載入。
- 填滿色彩在淺深色主題都一樣，所以有填滿的格子文字顏色依填滿決定（`inkFor`），不跟主題。
