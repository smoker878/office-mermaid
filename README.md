# Office Mermaid

沿用 Office Add-in Generator 的 TypeScript／Webpack 專案，以已安裝的 `pele@1.13.0` 渲染 Mermaid。使用 pnpm，不需後端。

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
- 目前以現代瀏覽器為目標，使用 Shadow DOM、Canvas、CSS variables；不宣稱支援 IE11，即使原始範本的 browserslist 尚包含 IE11。

## 驗證

```sh
pnpm typecheck
pnpm build:dev
pnpm build
```

啟動 `pnpm dev:browser` 後，開啟 `http://localhost:3000/tests.html`：無新增測試依賴的瀏覽器測試頁，完成後顯示每項結果，`document.body.dataset.testStatus` 為 `passed` 或 `failed`。測試只在 `--env tests` 時加入 bundle，不包含於一般 production build。

涵蓋主題解析、非法 CSS、flowchart／sequence／class／state／pie／subgraph、自包含 SVG、字型與色彩、無效／過長輸入、惡意 SVG 與頁面 CSS 隔離。

Production build 目前會提示 bundle 大小警告：完整 Pele 入口包含所有圖表類型，taskpane bundle 約 757 KiB（未 gzip）。這不是建置失敗；第一階段尚未進行按需載入最佳化。

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

> 目前只能在 Word 側邊欄使用 Mermaid／CSS 編輯器與 SVG 預覽。「插入圖表」、「載入選取圖表」、「更新圖表」仍停用；PowerPoint 尚未加入 manifest。以下是開發測試側載流程，不是正式發布安裝。

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

- 保留 `src`、`assets`、`tests`、`package.json`、`pnpm-lock.yaml`、`pnpm-workspace.yaml`、`manifest.xml` 與其他建置設定。
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
2. 到 **常用（Home）** 頁籤，尋找 **Show Task Pane** 按鈕；群組名稱仍可能顯示範本的 **Commands Group**。
3. 點擊按鈕，應顯示 Mermaid／CSS 編輯器與 SVG 預覽。
4. 檢查頁首是否顯示 Word 已連線，再修改 Mermaid 或主題顏色，確認預覽更新。
5. 此階段不會把圖表或原始碼寫入 Word 文件；停用的插入／載入／更新按鈕不是安裝失敗。

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
| 沒有 Show Task Pane 按鈕 | 查看 `pnpm start` 的側載錯誤，確認 Word 版本及組織是否禁止外掛。先保存錯誤訊息，不要自行修改登錄檔或降低信任中心安全設定。 |
| 想直接測 PowerPoint | 目前不支援；manifest 只有 Word，須待 PowerPoint 適配器與 manifest 整合後測試。 |

若仍失敗，請提供 Word 的版本／組建號、`pnpm start` 錯誤文字與側邊欄畫面；不要附上帳號、憑證或其他敏感資訊。

參考：[Microsoft：非本機伺服器測試與 start／stop 側載說明](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-non-local-server)。本專案尚未在 Windows Office 實測上述步驟。

## Office 階段：尚未完成

現有 `manifest.xml` 仍僅註冊 Word；`pnpm dev-server` 保留原本 HTTPS Office 開發流程，`taskpane.html` 非同步載入官方 Office.js，編輯器不等待 Office 初始化。瀏覽器模式請使用 `preview.html`。

Word／PowerPoint 插入、選取載入、文件內 metadata、更新原圖及存檔重開 **尚未實作、尚未驗收**。UI 中相關按鈕停用，不會把記憶體暫存假裝成文件內持久化。

下一階段先驗證 Word 的 SVG 與 Content Controls／文件內 metadata；之後另行驗證 PowerPoint 選取 Shape、ID、metadata 與替換 API。須取得實際宿主環境後，才能驗收保存重開、位置尺寸保留及複製副本獨立編輯。不得僅以模擬測試宣稱 Office 端到端支援。
