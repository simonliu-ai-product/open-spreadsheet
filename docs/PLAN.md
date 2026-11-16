# open-spreadsheet 技術規劃

類 Excel 的試算表框架，屬於 open-x「built for coding agents」系列。計算引擎和格子介面都自己寫，不用 Univer、HyperFormula 或其他試算表套件。目標是盡可能支援 xlsx／csv 會用到的公式。

> 狀態（2026-11-16）：M0–M2 完成，M3、M5 大部分完成。引擎 337 個函數，加上數字格式、日期時間輸入、公式複製位移、工作表改名／刪除；單元測試 206 個（含打包後從 node_modules 執行的檢查）、瀏覽器端對端檢查 113 項全部通過。介面風格與 open-doc／open-slide 一致（側欄、⌘K 指令面板、五種語言、淺深色）。有 `check`、skills、`current.json`、stdio MCP server（29 個工具，SHA-256 樂觀鎖）、儲存格備註與 `apply-notes`與 Claude Code／Antigravity plugin。npm：`@open-spreadsheet/core` 與短名稱 `open-spreadsheet`（`npx open-spreadsheet init`）0.4.1（單值函數對範圍逐格計算，同 Excel；窄寬度的編輯器頂部列；0.4.0 圖表風格、筆記插圖與 Markdown／PDF 匯出、素材與筆記的最近／最愛／搜尋、素材改名；0.3.0 Assets：圖表圖片與 Markdown 筆記；保護範圍；Excel 條件格式的前 N 名／平均／圖示集），`@open-spreadsheet/engine` 0.1.0。

---

## 1. 為什麼自己做

| 方案 | 結論 |
| --- | --- |
| Univer | 功能完整，但進階功能（匯入匯出、協作）要付費授權，而且整套 UI 框架很重，跟 open-x 「檔案就是原始碼」的方向不合 |
| HyperFormula | GPLv3 或商業授權，不適合 MIT 的 open-x |
| **自己寫** | 引擎是純 TypeScript、零依賴；資料格式由我們決定（agent 好讀好改、git 好 diff）。代價是公式要一個一個補，並用 Excel 的結果驗證 |

## 2. 架構

```
packages/
  engine/   @open-spreadsheet/engine  零依賴的計算引擎（Node 與瀏覽器共用）
  core/     @open-spreadsheet/core    CLI（dev / init）、Vite dev server、格子介面
  （之後）io/   xlsx／csv 匯入匯出，自己寫的 zip 讀寫（只需 deflate／inflate）
  （之後）mcp/  @open-spreadsheet/mcp  給 agent 的工具：讀範圍、寫儲存格、算公式
```

### 2.1 工作區格式

```
my-sheets/
  open-spreadsheet.config.json
  sheets/<id>/workbook.json      { title, sheets: [{ name, cells: { A1: "輸入值" }, cols? }] }
```

- 存的是**使用者輸入的原文**（`"120"`、`"=SUM(B2:B3)"`），不存計算結果。agent 直接改 JSON，dev server 監看檔案並推送 `open-spreadsheet:changed`，頁面重新載入。
- 頁面自己的寫入會被辨認出來，不會被當成外部修改。
- 之後加入格式（數字格式、粗體、框線、合併儲存格）時，放在 `styles` 欄位，不混進 `cells`。

### 2.2 引擎

- **值模型**：數字、文字、布林、錯誤、空白。轉型依 Excel 規則（`"1"+1=2`、`TRUE+1=2`、空白在加法裡是 0、在比較裡是 `""`）。
- **比較**：數字 < 文字 < 布林，文字不分大小寫，數字比較到 15 位有效數字。
- **日期**：1900 日期系統，保留 1900/2/29 這個 Lotus 相容錯誤，序號和 Excel 一致。
- **剖析器**：precedence climbing，負號優先於次方（`-2^2=4`），次方左結合（`2^3^2=64`），百分比、`&`、整欄整列參照（`A:A`、`1:1`）、跨工作表參照、`_xlfn.` 前綴、陣列常數。保留使用者的括號。
- **重算**：相依圖加拓撲排序，用明確的堆疊而不是遞迴，所以十萬格的長鏈也不會 stack overflow。循環參照回傳 `#CIRC!`（Excel 介面會顯示警告，我們先用錯誤值表示）。`NOW`、`TODAY` 等函數屬於 volatile，每次都重算。
- **插入／刪除列欄**：所有公式的參照跟著位移；被刪掉的參照變成 `#REF!`，範圍會縮小。

### 2.3 格子介面（MVP 已完成）

- 虛擬捲動：只畫看得到的列；欄寬可自訂。
- 一個永遠有焦點的 `<input>` 當編輯器，所以中文輸入法（IME）組字時不會掉字。
- 鍵盤：方向鍵、Shift 選範圍、Enter／Tab、F2、Delete、⌘Z／⇧⌘Z、⌘A。
- 複製貼上用 TSV，可以和 Excel、Numbers、Google 試算表互貼。
- 文字會溢到右邊空白的格子；數字靠右；錯誤用紅色。
- 公式列、名稱方塊、工作表分頁、自動存檔（350 ms debounce）。

## 3. 公式支援評估：對照微軟〈Excel 函數（依類別）〉

參考：<https://support.microsoft.com/zh-tw/excel/excel-functions-by-category>。微軟列了 15 類、約 527 個項目（同一個函數會出現在多個類別，例如新版清單和舊版相容清單，所以實際不重複的函數約 480 個）。

| 類別 | 微軟列出 | 已實作（約） | 能做到嗎 | 說明 |
| --- | ---: | ---: | --- | --- |
| 數學與三角函數 | 83 | 19 | ✅ 全部 | 純數學。`SUBTOTAL`、`AGGREGATE` 需要知道「被隱藏的列」，等介面支援隱藏列時一起做 |
| 統計 | 111 | 13 | ✅ 幾乎全部 | 需要自己寫分配函數（常態、t、F、卡方、Beta、Gamma），用標準數值方法即可。`FORECAST.ETS` 系列演算法微軟沒有完整公開，只能近似 |
| 文字 | 42 | 20 | ✅ 全部 | `TEXT` 的格式碼要完整支援；`ASC`／`JIS`／`PHONETIC` 這類全形半形與日文讀音函數可以做但優先度低 |
| 日期與時間 | 25 | 10 | ✅ 全部 | `NETWORKDAYS.INTL`、`WORKDAY.INTL`、`DATEDIF`、`YEARFRAC`（多種日數基準）都有明確規則 |
| 邏輯 | 19 | 11 | ✅ 全部 | `LET`、`LAMBDA`、`MAP`、`REDUCE`、`SCAN`、`BYROW`、`BYCOL`、`MAKEARRAY` 需要引擎支援「函數值」與動態陣列（第 3 階段） |
| 查閱與參照 | 40 | 11 | ✅ 大部分 | `FILTER`、`SORT`、`UNIQUE`、`XMATCH`、`TAKE`、`DROP`、`VSTACK` 等需要動態陣列溢出（spill）；`INDIRECT`、`OFFSET` 是動態參照，要讓相依圖支援。`RTD` ❌（需要 COM 伺服器）、`GETPIVOTDATA` 要等樞紐分析表 |
| 財務 | 55 | 4 | ✅ 全部 | 公式都有公開定義；`IRR`、`XIRR`、`RATE` 用牛頓法迭代，要對齊 Excel 的初始值與收斂條件 |
| 工程 | 54 | 0 | ✅ 全部 | 進位轉換、複數（`IMSUM` 等）、`BESSELJ`、`ERF`、`CONVERT`（單位表很長但是公開的） |
| 資訊 | 22 | 7 | ✅ 大部分 | `CELL`、`INFO` 只能回傳部分欄位（沒有檔案路徑、作業系統等資訊） |
| 資料庫 | 12 | 0 | ✅ 全部 | `DSUM`、`DGET` 等，條件區域的規則和 `SUMIFS` 類似 |
| 相容性 | 41 | 0 | ✅ 全部 | 大多是新版函數的舊名稱（`STDEV` → `STDEV.S`、`RANK` → `RANK.EQ`），做成別名即可 |
| Web | 3 | 0 | ⚠️ 部分 | `ENCODEURL` ✅；`WEBSERVICE`、`FILTERXML` 需要連網且有安全疑慮，預設停用，或交給 agent 處理 |
| Cube | 7 | 0 | ❌ | 需要連 OLAP／Power Pivot 資料模型，不在本專案範圍 |
| 使用者定義（增益集） | 3 | 0 | ❌ | `CALL`、`REGISTER.ID`、`EUROCONVERT` 依賴 Windows DLL 或增益集。改用 `LAMBDA`／具名函數取代 |
| 365／2024 新函數 | — | 部分 | ⚠️ 大部分 | `TEXTSPLIT`、`REGEXTEST` 等可以做；`STOCKHISTORY`、`TRANSLATE`、`DETECTLANGUAGE`、`IMAGE`、`COPILOT`、`PY` 依賴微軟雲端服務，❌ 或改接 agent |

**進度（2026-11-15）**：已實作 337 個。**結論**：約 480 個不重複函數裡，大約 95% 有機會做出來。做不到或不該做的大約 20 個：Cube 7 個、增益集 3 個、`RTD`、`WEBSERVICE`／`FILTERXML`（可選），以及依賴微軟雲端的 `STOCKHISTORY`、`TRANSLATE`、`DETECTLANGUAGE`、`COPILOT`、`PY`。真正花工夫的不是函數數量，而是三件引擎能力：**動態陣列溢出**、**LAMBDA 函數值**、**動態參照（INDIRECT／OFFSET）**。

## 4. 正確性驗證（oracle）

函數多了以後，不能靠手寫期望值。

1. **Excel 存檔時的快取值**：Excel 存 xlsx 時，每個公式格會附上 `<v>` 計算結果。準備一批「由 Excel 存檔」的 xlsx，引擎重算後和 `<v>` 逐格比較。這是最可靠的來源。
   - 注意：openpyxl 等工具產生的 xlsx 沒有快取值，不能用。
2. **LibreOffice headless** 作為第二意見（`soffice --convert-to`），差異時以 Excel 為準。
3. 微軟文件每個函數頁面的範例表格，整理成測試案例。
4. 數值比較用 15 位有效數字，錯誤值比較代碼。

## 5. 已知問題（spike 實測）

基準測試（`pnpm bench`，Apple Silicon）：

| 情境 | 建立加首次計算 | 改一個輸入 |
| --- | ---: | ---: |
| 10 萬格長鏈（`A2=A1+1`…） | 268 ms | 83 ms |
| 1 萬列加一個總和 | 34 ms | 5 ms |
| 2,000 列累計（`B2=SUM($A$1:A2)`） | 659 ms | 649 ms |

- ~~累計加總是 O(n²)~~ **已解決**：範圍讀者依欄分桶（`RangeIndex`，超過 64 欄的範圍另存一份清單），dirty 集合也能依欄查詢（`DirtySet`），整本載入時不逐格傳播。2,000 列累計加總：建立 659 → 194 ms、改第一格 649 → 162 ms；剩下的是 SUM 本身 O(n²) 的加法，Excel 也一樣。十萬格長鏈改第一格 83 → 114 ms，是索引維護的代價。
- 循環參照目前一律 `#CIRC!`，尚未支援 Excel 的「反覆運算」選項。

## 6. 里程碑

| 階段 | 內容 |
| --- | --- |
| **M0 spike**（完成） | 引擎核心、90 個函數、116 個測試、基準測試 |
| **M1 MVP**（完成） | CLI、dev server、格子介面、檔案監看、CoStaff Desktop 轉接描述檔、playground 範例 |
| **M2**（完成） | 欄寬列高拖曳與自動調整、插入／刪除列欄、填滿控點與 ⌘D／⌘R、數字格式（含自訂代碼）、凍結窗格、合併儲存格、文字樣式與顏色、排序、尋找取代、右鍵選單、工作表分頁操作、狀態列統計、縮放。範圍依欄建索引、dirty 依欄查詢（累計加總 649 → 162 ms） |
| M3（部分完成） | ✅ xlsx／csv 匯入匯出（自己寫 zip，含共用公式、格式、合併、凍結）。⬜ Excel 快取值 oracle 測試；✅ 函數補到 300 個以上（目前 337。第三批：工程類進位轉換、位元、複數、CONVERT，機率分配 NORM／T／CHISQ／F／BINOM／POISSON／GAMMA／BETA 與反函數，其餘數學統計；第二批：第二批補了數學三角、統計、文字、日期、XMATCH／LOOKUP／ADDRESS、資訊、財務、D 系列與相容性別名，測試用的是微軟文件裡的範例值） |
| M3b（完成） | 條件式格式（醒目提示、重複值、色階、資料橫條）、資料驗證下拉清單、自動篩選；範圍跟著插入／刪除列欄移動；`check` 與 MCP `set_layout` 都支援。✅ 這三項的 xlsx 匯入匯出（`conditionalFormatting`＋`dxfs`、`dataValidations`、`autoFilter`＋隱藏列；用 openpyxl 讀回驗證過結構） |
| M3c（完成） | 儲存格框線（所有、外框、粗外框、單邊，深色模式下黑線改用主題文字色）、手動隱藏／取消隱藏列欄（標題列在隱藏處畫藍線）；兩者都進出 xlsx（`borders`／`applyBorder`、`col hidden`、`row hidden`），MCP 與 skill 都支援 |
| M3d（完成） | 圖表：長條（分組）、折線、圓餅，手寫 SVG、無依賴；浮在格子上可拖曳、縮放、右鍵改類型／標題／範圍、⌫ 刪除、可復原；跟著篩選與隱藏列；Agent 用 MCP `set_layout` 的 `charts`。✅ 圖表的 xlsx 匯入匯出（DrawingML：`drawing*.xml` 錨點＋`chart*.xml`，數列直接參照儲存格，Excel 開啟後仍是活的；openpyxl 讀回驗證過） |
| M4 | 動態陣列溢出、`LET`／`LAMBDA`、`FILTER`／`SORT`／`UNIQUE`、`INDIRECT`／`OFFSET`；工程函數 |
| M5（部分完成） | ✅ 對齊 open-doc 近期功能：skills＋`sync-skills`、`current.json`、`check`、stdio MCP server（併在 core 的 `open-spreadsheet-mcp`，無外部依賴）、寫入鎖與 SHA-256 樂觀鎖、Claude Code／Antigravity plugin、⌘K、五種語言、經 proxy 的來源檢查。⬜ 發佈 npm（engine 需先建置成 JS：Node 不剝除 `node_modules` 內的 TS 型別）；⬜ 請 open-sheet 的維護者評估是否改用本專案 |
