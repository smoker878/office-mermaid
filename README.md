# Office Mermaid

由 Office Add-in Generator 的 TypeScript 專案改用 Vite，側邊欄使用 React 19 + Tailwind CSS 4，以 `pele@1.13.0` 渲染 Mermaid。使用 pnpm，不需後端。Vite 8 需要 Node.js 20.19+ 或 22.12+；建議沿用 Node.js 24 LTS。

## 前端架構

- `src/taskpane/taskpane.tsx`：React 入口，啟用 StrictMode；`App.tsx` 組合頁首、編輯器、預覽及文件操作列。
- `EditorPanel.tsx`／`PreviewPanel.tsx`：受控表單與 SVG 圖片預覽；保留原有藍灰配色、字級、4px 圓角與窄版排版。
- `use-editor.ts`：草稿、固定編輯目標、文件操作與選圖掃描；`use-diagram-preview.ts` 保留 400ms debounce、即時手動預覽及錯誤保留舊圖。
- `use-office-host.ts`：Office.js 初始化及事件訂閱／清除；`preview.html` 不載入 Office.js。
- `taskpane.css`：Tailwind 入口及 `@theme` 設計 tokens；共用 utility classes 位於 `styles.ts`。主題 CSS 與輸出 SVG 仍由原有渲染器隔離處理。
- `src/renderer/`、`src/office/` 與 `src/diagram/` 維持既有實作；文件資料格式、manifest 與 Office 側載流程不變。

Tailwind 在建置時產生一般 CSS，不需要 Office 安裝額外套件。Tailwind v4 需要 Chrome／Edge 111+、Safari 16.4+ 或 Firefox 128+；Windows 使用 WebView2，Mac 使用 WKWebView，仍須在目標 Office／作業系統實機驗證。不支援舊版 IE／EdgeHTML WebView。

### React／Tailwind 重構驗證

- `pnpm typecheck`、development／production build，以及 taskpane、相關測試與建置設定的 ESLint 通過；直接執行 ESLint，未執行含 usage data 的 Office CLI wrapper。
- Chromium：渲染測試 35／35、資料／Office／React 側邊欄模擬測試 35／35 通過。新增 debounce 寫入防護、錯誤恢復、重複插入鎖定、無效圖表載入保留草稿、失效更新目標及延遲事件訂閱清除測試。
- Production 預覽驗證：360px 無水平溢出、寬版雙欄、快捷鍵更新、CSS／排版欄位錯誤及恢復、舊預覽保留、SVG 解碼與非空白圖、離開提醒、封鎖網路後繼續編輯，以及瀏覽器模式不載入 Office.js／停用文件操作通過；production 不提供測試入口。
- 未執行實際 Word 側載、HTTPS 憑證、manifest 驗證或其他瀏覽器測試；Office 模擬測試不等同實機驗收。

## 打包與開發模式

- `pnpm dev:browser`：HTTP 瀏覽器預覽與測試頁，不讀取 Office 開發憑證。
- `pnpm dev-server`：`https://localhost:3000`，沿用 `office-addin-dev-certs`；首次使用可能要求建立／信任憑證。
- `pnpm start`／`pnpm stop`：保留原有 Office 側載流程；`start` 使用上述 `dev-server`。
- `pnpm build:dev`／`pnpm watch`：development 建置，保留 manifest 的 localhost 網址。
- `pnpm build`：production 建置；輸出的 manifest 沿用原本的 production 網址替換規則，設定集中在 `vite.config.mts` 的 `urlProd`（目前仍為範例網址，正式部署前須另行設定）。不修改原始 `manifest.xml`。
- 兩種開發模式固定使用 port 3000；已占用時直接失敗，不自動改 port，以免與 manifest 不一致。

根目錄的 `taskpane.html` 與 `commands.html` 是 Vite HTML 入口。`preview.html` 共用 taskpane 範本，瀏覽器網址維持不變以略過 Office.js。圖片仍位於 `/assets/`；建置後的 JS／CSS 使用帶 hash 的檔名。`tests.html` 與 `office-tests.html` 僅在 browser 開發模式提供，不加入一般建置。

Vite 不載入 `.env` 檔；不再使用 Webpack、Babel 或 IE11 polyfill。請使用 pnpm 與 `pnpm-lock.yaml`，不要使用原範本遺留、早已不同步的 `package-lock.json` 安裝依賴。

### Vite 遷移驗證

- `pnpm typecheck`、development／production build，以及本次修改的 TypeScript／Vite 設定 ESLint 通過。
- Chromium：35／35 渲染測試、29／29 資料／Office／側邊欄模擬測試通過。
- 驗證兩種建置的 HTML／圖片產出、preview 共用內容、manifest 網址替換、測試入口排除，以及原始 manifest 未變更。
- Production 靜態預覽：SVG 解碼、修改後重新預覽、錯誤保留舊預覽、瀏覽器模式不載入 Office.js／停用插入、360px 無溢出通過；測試頁回傳 404。
- Production 主 bundle 約 728 kB（gzip 約 269 kB），仍有大型 chunk 警告；未做渲染器按需載入。
- 未測試 Office HTTPS 憑證、實際 Word 側載、manifest 驗證與其他瀏覽器；不宣稱已完成 Office 實機驗收。

## 第一階段：瀏覽器編輯器

```sh
pnpm dev:browser
```

開啟 `http://localhost:3000/preview.html`。此模式使用 HTTP，不讀取或安裝 Office 開發憑證，也不需要 Office、Microsoft 365 帳號或 Office.js CDN。

- Mermaid 與 CSS 多行編輯器；輸入後延遲 400ms 更新預覽。
- 「更新預覽」或 Ctrl／⌘ + Enter 可立即渲染。
- 錯誤時保留輸入與上一次成功預覽，並指出錯誤欄位；舊預覽不代表目前輸入有效。
- 輸入尚未寫入文件或瀏覽器儲存；離開時有未保存提醒，但瀏覽器可能限制提醒。
- 頁面載入後可斷網繼續編輯；不保證斷網時能首次開啟或重新載入頁面。

### 調整大小：排版與縮放分開

- **可用寬度（200–4000 px）**：輸入數字或拖動滑桿，重新產生 SVG。支援的圖表可能換行、調整間距；不是單純縮放。
- **依寬度自動調整方向**：預設關閉以保留原始方向；開啟後，較窄且放不下的圖表可能轉成直向。不是每種圖表都支援，實際輸出可能超過指定寬度。
- **圖片寬度（1–50 公分）**：設定 Word 內的顯示寬度，高度依 SVG 比例計算，不改排版。預設 14 公分，請自行確認頁面、欄位及表格可用空間。
- 更新既有圖表預設讀取並保留 Word 中的目前寬度；取消勾選後，才使用輸入的新寬度。排版設定與原始碼一併存入文件。
- 預覽會縮放以符合側邊欄，但側邊欄寬度不會改變產生的 SVG 排版。

### CSS 範圍

只接受單一 `.pele` 區塊與支援的 `--pele-*` 變數：

```css
.pele {
  --pele-bg: #ffffff;
  --pele-fg: #243247;
  --pele-surface: #edf3fc;
  --pele-accent: #2563eb;
  --pele-font: sans-serif;
  --pele-radius: 4px;
}
```

顏色限 HEX（3／4／6／8 位）；支援 bg、fg、fg-on-light、fg-on-dark、muted、line、surface、surface-alt、border、accent、critical 與 series-1 至 series-8。

其他變數：font、font-mono（本機字型與 fallback 清單）、radius（0–64px）、title-weight／heading-weight／bold-weight（normal、bold、100–900）。不接受任意 selector、外部資源、@ 規則、`!important` 或巢狀 `var()`。Mermaid 本身的固定樣式可能優先於主題。

### SVG 與限制

渲染與 Office API 分離於 `src/renderer/`。SVG 先進行元素與屬性白名單驗證，再於隔離的 Shadow DOM 中解析樣式，轉為固定的 SVG presentation attributes。預覽直接使用這份自包含 SVG，不依賴編輯頁面的 CSS。

- 禁止 script、事件處理器、外部圖片、連結、foreignObject、外部 CSS 與 XML 實體。
- Mermaid 限 50,000 字元、1,000 條邊；CSS 限 10,000 字元；SVG 限 4,000,000 字元與每邊 30,000px。大型圖仍可能短暫阻塞主執行緒。
- 不內嵌字型；不同裝置字型與字寬可能不同。尚未驗證 Office 對輸出 SVG 的相容性。
- 目前以現代瀏覽器／Office WebView 為目標，使用 Shadow DOM、Canvas、CSS variables；不支援 IE11 或舊版 Office WebView。實際宿主相容性仍須實測。

## 驗證

```sh
pnpm typecheck
pnpm build:dev
pnpm build
```

啟動 `pnpm dev:browser` 後，開啟 `http://localhost:3000/tests.html`：無新增測試依賴的瀏覽器測試頁，完成後顯示每項結果，`document.body.dataset.testStatus` 為 `passed` 或 `failed`。測試只在 browser 開發模式提供，不包含於一般 production build。

涵蓋主題解析、非法 CSS、flowchart／sequence／class／state／pie／subgraph、自包含 SVG、字型與色彩、無效／過長輸入、惡意 SVG 與頁面 CSS 隔離。

Production build 目前會提示 bundle 大小警告：完整 Pele 入口包含所有圖表類型，加上 React 後 taskpane bundle 約 958 kB（gzip 約 341 kB）。這不是建置失敗；尚未進行按需載入最佳化。

另以編輯頁人工／瀏覽器自動化驗證：

1. 修改 Mermaid，確認即時預覽與尺寸更新。
2. 輸入錯誤語法、清空 Mermaid，確認提示、輸入保留、舊預覽不被覆蓋。
3. 修改 CSS 顏色，確認主題生效；輸入非法 CSS 後修正，確認恢復正常。
4. 測試 Ctrl／⌘ + Enter、鍵盤焦點及 360px 窄版無水平溢出。
5. 載入頁面後切換離線，再修改圖表；確認仍能渲染。

### 本次驗證紀錄

- `pnpm typecheck`、development build、production build 通過。
- 使用現有 `office-addin-lint` 所附 ESLint 設定，直接檢查 5 個新增／修改的 TypeScript 檔案，通過（未執行會回傳 usage data 的 CLI wrapper）。
- Chromium 瀏覽器測試頁：35／35 通過；外部資源攔截記錄為 0 筆請求。
- 編輯頁自動化：8 項互動檢查通過，包括錯誤保留預覽、欄位標記與恢復、主題生效、空內容引導、快捷鍵與 Office 按鈕停用。
- 360px 無水平溢出；頁面載入後離線仍能修改／渲染；實際 SVG 圖片解碼及 Canvas 非背景像素檢查通過，非空白圖。
- 未執行 Office 宿主測試、manifest 驗證或其他瀏覽器相容性測試；未安裝依賴、部署或修改原始 manifest。

## 在另一台 Windows 的 Word 開啟外掛

> 已加入 Word 的插入、載入與更新程式，須支援 WordApi 1.4；尚未經實際 Word 驗收。PowerPoint 尚未加入 manifest。以下是開發測試側載流程，不是正式發布安裝。請先使用文件副本測試。

### 1. 準備 Windows 環境

- 安裝並啟用 Word 桌面版；建議先使用更新到最新版本的 Microsoft 365 Word。
- 在 **Word → 檔案 → 帳戶 → 關於 Word** 記下產品名稱、版本與組建號，供後續 API 相容性驗證。其他 Office 版本不保證支援相同功能。
- 安裝 [Node.js 24 LTS](https://nodejs.org/)。
- 安裝 pnpm；本專案開發時使用 `12.3.4`。若尚未安裝，可在 PowerShell 執行以下命令（會下載並全域安裝 pnpm）：

  ```powershell
  npm install --global pnpm@12.3.4
  ```

重新開啟 PowerShell，確認命令可用：

```powershell
node --version
pnpm --version
```

若電腦由公司管理，先確認允許安裝開發工具、信任本機開發憑證與側載 Office Add-in；不要繞過組織的安全政策。

### 2. 將專案複製到 Windows

將專案資料夾複製到 Windows，例如 `C:\Projects\office-mermaid`：

- 保留 `src`、`assets`、`tests`、根目錄 HTML 入口、`vite.config.mts`、`package.json`、`pnpm-lock.yaml`、`pnpm-workspace.yaml`、`manifest.xml` 與其他建置設定。
- 不要複製 `node_modules`、`dist` 或 Mac 的開發憑證；Windows 上重新安裝依賴及建立憑證。
- 不需要重新執行 Office Add-in Generator，也不要另建專案。

目前 manifest 使用 `https://localhost:3000`。**localhost 指的是執行 Word 的 Windows 電腦，不是原本開發的 Mac**，所以只在 Mac 啟動伺服器、將 manifest 複製到 Windows 並不足夠。跨電腦連線需要另外設定可存取的 HTTPS 位址、憑證與 manifest，不屬於本流程。

### 3. 安裝依賴並啟動 Word 測試

在 Windows 專案根目錄（包含 `package.json` 的資料夾）開啟 PowerShell：

```powershell
Set-Location "C:\Projects\office-mermaid"
pnpm install --frozen-lockfile
pnpm start
```

`pnpm install` 會下載依賴，並依專案設定執行允許的安裝腳本。若 lockfile 不相容，先記下錯誤與 pnpm 版本，不要刪除 lockfile 或改用 npm install。

現有 `start` 指令執行 `office-addin-debugging start manifest.xml`，設定目標為 Word 桌面版，會嘗試啟動 HTTPS 開發伺服器、註冊測試外掛並開啟 Word。

首次啟動可能要求安裝／信任本機開發憑證；確認是本專案的 `office-addin-dev-certs` 提示後再允許。**不要停用憑證驗證或信任來源不明的憑證。** 測試期間保留開發伺服器執行；首次載入 Office.js 需要能連線至 Microsoft CDN。

> Word 測試使用 `pnpm start`。`pnpm dev:browser` 是 HTTP 瀏覽器預覽模式，不可直接替代 manifest 指向的 HTTPS 服務。同一台電腦若已執行瀏覽器預覽，先以 Ctrl+C 停止，避免占用相同的 3000 port。

### 4. 在 Word 開啟側邊欄並檢查

1. 若 Word 尚未開啟文件，建立一份空白文件。
2. 到 **常用（Home）** 頁籤，尋找 **Mermaid** 群組的 **Mermaid 圖表** 按鈕。若仍顯示舊名稱 Show Task Pane，需重新載入更新後的 manifest。
3. 點擊按鈕，應顯示 Mermaid／CSS 編輯器與 SVG 預覽。
4. 檢查頁首是否顯示 Word 已連線，再修改 Mermaid 或主題顏色，確認預覽更新。
5. 將游標放在空白位置，按「插入文件」；成功後選取該圖表，按頂部「編輯選取圖表」，修改後按底部「更新這張圖表」。詳細驗收流程見下節。若 API 不支援，頁首會說明原因。

### 5. 結束測試

在同一個專案目錄執行：

```powershell
pnpm stop
```

此命令會嘗試停止偵錯服務並移除側載的測試註冊；不要只關閉 Word 或終端機來代替。

### 常見問題

| 情況 | 先檢查 |
| --- | --- |
| 找不到 `node` 或 `pnpm` | 安裝後重新開啟 PowerShell，確認工具已加入 PATH。若 PowerShell 阻擋 `.ps1`，可改用命令提示字元執行相同 pnpm 指令，不要直接放寬整台電腦的執行原則。 |
| 3000 port 已被占用 | 停止本專案先前啟動的預覽／開發服務；不要任意結束不明程序。 |
| 側邊欄空白或 HTTPS 錯誤 | 在 Windows 瀏覽器開啟 `https://localhost:3000/taskpane.html`，先確認服務可連線且沒有憑證警告；瀏覽器可顯示頁面不代表 Office API 已通過驗證。 |
| 沒有 Mermaid 圖表按鈕 | 查看 `pnpm start` 的側載錯誤，確認 Word 版本及組織是否禁止外掛。先保存錯誤訊息，不要自行修改登錄檔或降低信任中心安全設定。 |
| 想直接測 PowerPoint | 目前不支援；manifest 只有 Word，須待 PowerPoint 適配器與 manifest 整合後測試。 |

若仍失敗，請提供 Word 的版本／組建號、`pnpm start` 錯誤文字與側邊欄畫面；不要附上帳號、憑證或其他敏感資訊。

參考：[Microsoft：非本機伺服器測試與 start／stop 側載說明](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-non-local-server)。本專案尚未在 Windows Office 實測上述步驟。

## Word 文件插入與重新編輯（待實機驗收）

現有 `manifest.xml` 仍僅註冊 Word；`pnpm dev-server` 保留原本 HTTPS Office 開發流程，`taskpane.html` 非同步載入官方 Office.js，編輯器不等待 Office 初始化。瀏覽器模式請使用 `preview.html`。

### 使用方式

1. 新增：編輯 Mermaid，按需展開「樣式與 CSS」與「排版與圖片大小」。將 Word 游標放到內容控制項以外的空白插入位置，按底部「插入文件」。選取文字或圖片時會拒絕，避免誤刪內容。
2. 編輯：在 Word 選取一張外掛圖表，或把游標放入其容器，按頂部「編輯選取圖表」。選圖提示出現時，也可按「編輯這張」。選取本身不會取代草稿；編輯器有未寫入的變更時，切換前會先詢問。
3. 更新：修改後按底部「更新這張圖表」。頂部會顯示已載入圖表的原始碼摘要。目標為先前載入的圖表，不會隨游標移動；預設保留目前文件圖片寬度，高度依比例重新計算。
4. 新增另一張：按頂部「＋ 新增圖表」，保留目前編輯內容，再把 Word 游標移到新位置插入。這不是清空草稿。
5. 最後在 Word 儲存 `.docx` 文件。外掛顯示「已寫入目前文件」不等於已存檔到磁碟。

雙擊文件圖片不會開啟側邊欄，請使用 Word 功能區的「Mermaid 圖表」。選圖監聽失敗時會提示改用手動編輯入口；一般圖片、空白與多選不會自動切換目標或跳出錯誤，手動載入時才顯示具體原因。

本次介面調整的本地驗證：typecheck、針對修改 TypeScript 的 ESLint、production build、35／35 渲染測試，以及 29／29 資料／Office／側邊欄模擬測試；涵蓋選圖提示、草稿保護、固定更新目標、空白／多選、過時提示與延遲掃描、進階設定錯誤展開與 360px 無水平溢出。選取事件與新版 manifest 仍需 Word 實機驗收；production build 仍有 bundle 大小警告。

### 儲存與安全策略

- `src/office/word-adapter.ts` 使用 Rich Text Content Control 包住行內圖表，tag 指向文件內 Custom XML Part；保存 Mermaid、CSS、排版、顯示寬度、圖表 ID 與版本。
- `src/office/ooxml.ts` 產生 Flat OPC，內嵌 SVG 與由同一份 SVG 產生的 PNG fallback，不連結外掛頁面或外部圖片。Word 是否採用 SVG、顯示一致性仍須實測。
- 每次更新新增不可變的原始碼快照，不改寫舊快照。相同文件內複製的容器先共享快照，更新其中一份時會分離來源關聯；只有目標容器會被替換。
- 不自動清除歷史／未引用快照，以利復原及保護複製副本。文件大小會隨更新增加；歷史快照可能保留已刪除的原始碼，不適合當作敏感內容清除機制。
- Office 批次不是交易；更新前保留原圖 OOXML，失敗時嘗試還原。無法確認還原時會提示使用 Word 復原並先檢查文件，不自動重試。
- 渲染、資料或尺寸驗證失敗不更新圖表；寫入前驗證錯誤保留編輯目標。目標遺失、衝突或寫入結果不確定時，保留草稿但要求重新載入確認。

### 已知限制

- 需要執行期檢查支援 `WordApi 1.4`；不代表所有具備此 API 的 Office 組建都已驗證 SVG 相容性。
- 第一版只管理單一行內圖片；不支援浮動圖、旋轉／裁切保留，或在容器加入其他文字／巢狀內容控制項後繼續更新。更新保留寬度，不保證完全保留原圖所有格式。
- 追蹤修訂開啟時拒絕插入與更新。鎖定、唯讀、受保護文件可能拒絕寫入；請保留編輯內容並解除限制後重試。
- 不支援跨文件複製原始碼關聯；跨文件貼上的圖片可能可見，但不保證其 Custom XML Part 一起複製，遺失時會清楚提示。
- 共編／外部修改有更新前的目標與來源檢查，但不是跨使用者鎖定或原子交易；第一版請避免同時編輯同一圖表。
- PowerPoint 尚未實作；manifest 維持 Word，沒有變更部署或憑證設定。

### 無 Word 的本地測試

啟動 `pnpm dev:browser` 後：

- `http://localhost:3000/preview.html`：編輯、尺寸與預覽；文件操作停用。
- `http://localhost:3000/tests.html`：既有渲染與 SVG 安全測試。
- `http://localhost:3000/office-tests.html`：資料往返、排版、PNG／SVG OOXML 與 Word API 模擬測試。模擬的載入測試只表示讀取文件模型，不是實際存檔重開驗證。

本輪本地驗證：TypeScript 檢查、針對本次 7 個 TypeScript 檔案的 ESLint、production build 通過；35／35 渲染測試與 21／21 資料／Office 模擬測試通過。另以 Chromium 驗證排版變更、尺寸錯誤後恢復、舊預覽保留、瀏覽器模式按鈕停用、360px 無溢出，以及模擬宿主的載入／更新、驗證失敗後重試與 debounce 載入失敗後恢復。Production bundle 約 786 KiB（未 gzip），仍有大小警告。未執行 manifest 驗證、Office 宿主或其他瀏覽器測試。

### 實際 Word 驗收清單

請使用 `.docx` 副本，記錄作業系統、Word 版本／組建號及每項結果：

- [ ] 插入含中文與 CSS 的 SVG；沒有外掛時圖片仍可見。
- [ ] 同一份文件放兩張圖，載入其中一張，移動游標後更新，確認只改原目標。
- [ ] 調整排版寬度並開關自動方向，確認插入圖與預覽一致。
- [ ] 在 Word 手動縮放圖片後更新，確認保留寬度；取消保留後使用新寬度且不變形。
- [ ] 儲存、關閉文件與側邊欄，重新開啟後選取圖表，還原 Mermaid、CSS 與排版設定並更新。
- [ ] 在同文件複製貼上一張圖，分別更新兩份，確認不互相影響。
- [ ] 測試一般圖片、未選取、多選、原圖被刪除、唯讀／鎖定與追蹤修訂的提示。
- [ ] 語法錯誤時原圖不變，修正後能繼續更新；測試 Word 復原／重做後的圖表與原始碼關聯。
- [ ] 如遇插入結果不確定或還原失敗，依提示檢查文件，不直接重複操作。

本機沒有 Word，以上實機項目尚未執行，不宣稱 Word 端到端驗收完成。
