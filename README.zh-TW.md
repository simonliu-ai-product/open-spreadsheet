# open-spreadsheet

**為 Coding Agent 打造、用起來像 Excel 的試算表。** 每份活頁簿都是一個 JSON 檔，存的是人輸入的內容——`"120"`、`"2026-01-05"`、`"=SUM(B2:B9)"`——人在真正的格子裡編輯，Agent 直接改檔案，git 看得到兩邊的差異。

[English](README.md)

## 特色

### 🧮 自己寫的公式引擎

底層沒有任何試算表套件。引擎（`packages/engine`，零依賴）在最容易出相容問題的地方都照 Excel 的規則：型別轉換、15 位有效數字比較、含閏年錯誤的 1900 日期系統、`-2^2 = 4`、跨工作表範圍、整欄（`A:A`）、絕對參照，以及插入／刪除列欄時改寫所有公式（被刪掉的儲存格變成 `#REF!`）。重算以相依圖加上明確堆疊的拓撲排序完成，十萬格的長鏈也不會溢位。目前 337 個函數——數學與三角、統計（MAXIFS、PERCENTILE、RANK、CORREL、FORECAST）、文字（TEXTBEFORE／AFTER、TEXT、FIXED）、日期（NETWORKDAYS、WORKDAY、DATEDIF、YEARFRAC）、查閱（XLOOKUP、XMATCH、INDEX/MATCH）、財務（PMT、RATE、IRR、XIRR、NPV）、工程（進位轉換、位元運算、複數、CONVERT）、機率分配（NORM、T、CHISQ、F、BINOM、POISSON、GAMMA、BETA）、資料庫 D 系列函數與舊版名稱；對照微軟函數清單的完整規劃見 [docs/PLAN.md](docs/PLAN.md)。

### 🟩 行為像 Excel 的格子

- 公式列與名稱方塊（輸入 `B12` 或 `A1:C5` 直接跳過去）、函數自動完成與參數提示、輸入公式時以顏色框出參照、用滑鼠或方向鍵點選參照、`F4` 切換 `$`。
- 數值格式（`#,##0`、`"NT$"#,##0`、`0.0%`、`yyyy-mm-dd`、含多區段與 `[Red]` 的自訂格式）；輸入 `2026/1/5`、`12%`、`$1,200` 會自動套上對應格式。
- 粗體、斜體、底線、刪除線、文字與填滿色彩、框線（所有、外框、粗外框、單邊）、對齊、自動換行、合併儲存格、凍結窗格、隱藏列欄、拖曳調整欄寬列高（雙擊自動調整）。
- 儲存格、列欄標題與工作表分頁都有右鍵選單：剪下／複製／貼上（含只貼上值）、插入與刪除列欄、清除、排序、合併、凍結。
- 工作表分頁：新增、重新命名（雙擊）、建立副本、刪除、拖曳排序、標籤顏色。
- 定義名稱（`=SUM(業績)`）：從選取範圍定義、名稱管理員、名稱方塊輸入名稱即跳到該範圍、公式自動完成會列出名稱；名稱會跟著插入列與工作表改名調整，也會隨 .xlsx 進出。
- 追蹤前導參照／從屬參照（同 Excel）：虛線框和箭號標出公式讀了哪些格子，或哪些公式讀了這一格，名稱會展開；清單列出每一個（含其他工作表）並可點了跳過去。按 Esc 或一編輯就移除箭號。
- Excel 檔案的條件格式：除了比較、文字、重複值、色階、資料橫條，也支援前／後 N 名、高於／低於平均與圖示集（箭頭、燈號、符號、旗幟、評等），會顯示在格子裡，匯出 .xlsx 也保留。
- 保護範圍：從右鍵選單或 ⌘K 鎖住儲存格；輸入、貼上、清除、填滿、排序、改格式都會被拒絕，MCP 寫入也一樣，agent 改不到人鎖住的地方。
- 核取方塊（值為 TRUE／FALSE）：從右鍵選單或 ⌘K 插入，點一下勾選，空白鍵切換整個選取範圍；`=COUNTIF(B2:B9, TRUE)` 這類公式可以直接計算。
- 超連結：輸入網址或信箱，或用 `=HYPERLINK(網址, 名稱)`，就會顯示成連結；⌘＋點擊或右鍵「開啟連結」開啟（只接受 http、https、mailto）。
- 垂直對齊（靠上／置中／靠下），從右鍵選單或 ⌘K 設定，匯出 .xlsx 與列印都保留。
- 字型大小（同 Excel 的點數，預設 11）：工具列選單與 ⇧⌘> / ⇧⌘<；沒手動調過高度的列會自動長高；匯出 .xlsx 與列印都保留。
- 自訂排序：最多三個欄位，各自遞增或遞減，可選第一列是否為標題；同值的列維持原本順序。
- 顯示公式（⌃`）讓每一格顯示公式而不是結果；「跳到下一個錯誤」逐一走過活頁簿裡的錯誤格，並說明錯誤的意思（#DIV/0!、#REF!、#NAME? 等）。
- 資料工具：移除重複項、資料剖析（依分隔符號拆欄）、轉置貼上。
- 依同欄已有內容自動完成輸入、複製格式（⇧⌘C／⇧⌘V 或刷子按鈕）、填滿控點（數列、日期、`項目 1 → 項目 2`、參照會位移的公式）、`⌘D`／`⌘R`、`⌘Enter` 填滿選取範圍、與 Excel 和 Google 試算表互相複製貼上（TSV）、尋找與取代、復原與重做、縮放，以及顯示加總、平均、最小、最大、計數的狀態列。
- **Assets**：每本活頁簿旁邊的 `sheets/<id>/assets/` 與共用的 `assets/`，放圖表圖片（長條、折線、圓餅，用工具列的圖表按鈕從選取範圍畫成 SVG，不再放在格子上）、上傳的圖片和 Markdown 筆記。筆記記下數字的意義與規則，agent 處理前會先讀。儲存格備註寫「畫一張每月業績長條圖」，經 `apply-notes` 就會在 Assets 得到圖片。舊活頁簿格子上的圖表可以一鍵移到 Assets。圖表風格（六種內建，自訂的存在 `assets/chart-styles.json`）決定顏色、字型與圓角；筆記可用 `![標題](chart.svg)` 顯示圖表，並匯出成 Markdown 或 PDF；素材與筆記都有最近、最愛與搜尋。
- 條件式格式（大於／小於／介於／等於、文字包含、重複值、色階、資料橫條）、只能從清單選值的下拉選單，以及表頭附下拉選單的篩選——都存在活頁簿檔案裡。
- 儲存格備註（右鍵或 ⇧F2）：角落標記、滑過顯示、標題列清單——也可以當成給 Agent 的修改要求（`apply-notes`）。
- ⌘K 指令面板；淺色與深色；English、繁體中文、简体中文、日本語、한국어。

### 🖨️ 列印或存成 PDF

⌘P（或「下載 → 列印」）把工作表排成分頁表格：保留數值格式、樣式、合併儲存格與框線，略過隱藏列欄，凍結的標題列每頁重複。系統的列印對話框可以直接存成 PDF。

### 📥 CSV 與 Excel 進出

匯入 `.csv`、`.tsv`、`.xlsx`（工作表、公式——含共用公式、數值格式、字型、填滿、欄寬、合併、凍結窗格），下載 `.xlsx` 或 CSV。zip 讀寫是自己寫的，檔案不會上傳到任何地方。

### 📊 餵資料給 open-dashboard

`open-spreadsheet export <活頁簿> [--sheet] [--range A1:F200] [--format csv|json] [--out 檔案]` 輸出工作表算好的值——公式已計算、數字不加千分位、日期為 ISO 格式、第一列為表頭。[open-dashboard](../open-dashboard) 的收集器執行它、`csv` 來源讀取輸出檔，儀表板就能畫出大家在格子裡維護的資料：

```ts
// open-dashboard.config.ts
datasources: { budget: { type: 'csv', tables: { orders: { file: 'data/orders.csv' } } } },
collectors: {
  orders: { run: 'npx -p @open-spreadsheet/core open-spreadsheet export sales-report --range A1:F200 --out data/orders.csv --root ../my-sheets', source: 'budget', every: '5m' },
},
```

### 🤖 為 Agent 設計

- `init` 會寫入 `AGENTS.md`、`CLAUDE.md`、`.mcp.json` 與四個 Skills（`create-sheet`、`sheet-authoring`、`current-sheet`、`apply-notes`）；`sync-skills` 用來更新。
- 編輯器把目前選取的位置記在 `node_modules/.open-spreadsheet/current.json`，Agent 聽得懂「這一格」。
- `open-spreadsheet check` 列出顯示 `#NAME?`、`#REF!`、`#CIRC!`、`#VALUE!` 的公式，以及格式錯誤的儲存格、樣式與合併範圍。
- `open-spreadsheet-mcp` 是透過 stdio 的 MCP server，29 個工具——`read_range`、`trace_cell`、`list_assets`、`read_asset`、`write_note`、`render_chart`、`list_chart_styles`、`save_chart_style`、`rename_asset`、`delete_asset`、`write_cells`、`set_style`、`set_layout`、`create_workbook`、`rename_workbook`、`duplicate_workbook`、`delete_workbook`（移到 `sheets/.trash/`，不會真的刪掉）、工作表與列欄操作、`list_notes`、`set_note`、`set_names`、`evaluate`、`check`、`get_current`。每次寫入都會重算、回報留下的問題，而且檔案在讀取後被改過就拒絕寫入（SHA-256 檢查）。檢視器存檔走同一套規則，Agent 和人不會互相覆蓋——Agent 一改，開著的格子馬上更新。
- Claude Code 與 Antigravity 的 plugin 在 [plugins/open-spreadsheet](plugins/open-spreadsheet)：

```bash
# Claude Code
/plugin install open-spreadsheet --marketplace simonliu-ai-product/open-spreadsheet

# Google Antigravity
agy plugin install https://github.com/simonliu-ai-product/open-spreadsheet
```

## 快速開始

```bash
npx -p @open-spreadsheet/core open-spreadsheet init my-sheets
cd my-sheets && pnpm install && pnpm dev    # http://localhost:5373
```

也可以用 [CoStaff Desktop](../open-desktop) 開啟資料夾。

## 檔案契約

見 [README.md](README.md#the-file-contract) 與 `sheet-authoring` skill。

## 開發

```bash
pnpm test        # 引擎、ops 與 MCP 測試（vitest）
pnpm typecheck
pnpm bench       # 重算效能
node scripts/ui-e2e.mjs <輸出資料夾>   # 用 Chromium 操作格子並截圖
```

## 授權

MIT
