# Word Markdown Companion Add-in

## 專案目標

提供一個 Word 任務窗格增益集，讓使用者可以：

- 匯入 `.md` 檔案內容並插入到目前 Word 文件。
- 將目前 Word 文件內容輸出為 Markdown 字串（可複製或另存）。

此版本是「最小可用版本」(MVP)：

- 不包含完整的離線儲存/版本控管。
- 匯出階段使用純文字轉 Markdown，未做完整的樣式反向映射（標題與清單有基礎轉換，但無法保證 100% 等價）。
- 主要目的是快速建立「Word + Markdown」的載入/互動流程，作為後續加強基礎。

## 目錄結構

- `manifest.xml`：Office Add-in manifest，提供 Word 任務窗格入口。
- `src/taskpane.html`：任務窗格主頁。
- `src/js/taskpane.js`：匯入/匯出邏輯。
- `src/styles/taskpane.css`：任務窗格樣式。
- `assets/icon-16.png`、`assets/icon-32.png`、`assets/icon-80.png`：add-in 圖示。
- `scripts/dev-server.js`：本地靜態伺服器。
- `docs/publish-online.md`：線上可上架版本的 manifest 與發布說明。
- `src/lib/marked.min.js`、`src/lib/turndown.min.js`：本地 Markdown 轉換腳本。
- `.clean/legacy/`：已停用但暫時保留的舊流程檔案。
- `package.json`：專案指令。

## 快速開始

1. 安裝 Node.js 20+。
2. 在專案根目錄執行一個指令（自動安裝＋產生 manifest＋啟動本機伺服器）：
   ```bash
   npm run setup
   ```
   服務會使用 `MANIFEST_HOST`（預設 `http://localhost:3000`）產生 `manifest.xml`，並啟動 `dev-server`。
3. 若你要自動 sideload 到桌面版 Word，請在另一個終端機執行：
   ```bash
   npm run sideload
   ```
   這個指令會重新產生 `manifest.xml`、啟動本機伺服器，並將 manifest 註冊到 Office 的 sideload registry，供 Word Desktop 重新開啟後載入。

如果你偏好手動 sideload，仍可在 Word 選擇「上傳我的清單」/Upload My Add-in，載入 `manifest.xml`。

如果你要把 `.md` 檔關聯到 Word（可選），可用一鍵同時完成：

```bash
npm run setup:with-md-association
```

如果你要產生線上可上架版本的 manifest：

```powershell
$env:MANIFEST_HOST = "https://your-addin-host.example"
$env:SUPPORT_URL = "https://your-addin-host.example/support"
npm run build:store
```

輸出檔案會在 `dist/manifest.store.xml`。

#### 進階啟動

- 本地化指令（保留舊流程）：
  - `npm run start:local`
  - `npm run start:local -- --host https://addin.example.internal`
- 桌面版 Word 自動 sideload：
  - `npm run sideload`
  - 建議先關閉 Word 再執行，成功率較高。
  - 若 Word 顯示無法從 localhost 載入 add-in，通常還需要替 WebView 開啟 loopback exemption。
- 分步驟舊流程：
  - `npm install`
  - `npm run dev`

## 使用方式

1. 在 Word 開啟文件後，從 add-in 按鈕開啟任務窗格。
2. 點「插入 Markdown 到文件」：
   - 選擇 `.md` 檔案。
   - 檔案內容會被轉為 HTML 插入到游標位置。
3. 若你用系統關聯雙擊 `.md` 檔，launcher 會先保存 Markdown 內容、再開啟空白 Word 視窗。taskpane 會優先偵測這個待匯入內容，必要時自動匯入，或顯示「匯入剛開啟的 Markdown 檔」按鈕。
4. 點「匯出為 Markdown」：
   - 點選後會把目前 Word 文件轉為 Markdown 字串顯示在文字區。
   - 點「下載為 Markdown 檔」可直接儲存成 `.md`。

## 注意事項

- 這是 Word 任務窗格層級的「導入／導出」整合，能讓使用者在 Word 內直接操作 `.md`。  
- 不能直接把 `.md` 當成 Word 文件交給 `WINWORD.EXE "%1"` 打開。Word 會把它當文字檔/受保護內容處理，Office add-in 在這種模式下容易被鎖住或無法正常互動。
- 目前專案改用 launcher-based flow：雙擊 `.md` 先把內容交接給本機橋接，再開啟空白 Word 視窗，由 add-in 匯入 Markdown。
- Office Add-in 開發建議使用 `https://localhost`。目前這個專案仍以 `http://localhost:3000` 為主，某些 Word/Office 環境可能會直接拒絕或顯示不安全內容警告。
- 若桌面版 Word 出現「We can't open this add-in from localhost」，請以系統管理員權限執行：
  ```powershell
  CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
  ```

## 目前狀態更新

- `marked` 與 `turndown` 仍以 `src/lib` 本地載入，降低 Markdown 轉換階段對外部 CDN 的依賴。
- `office.js` 已改回 Microsoft 官方 hosted 版本。原因是 Office runtime 會動態載入 host-specific 與 locale 腳本，單獨 vendoring `office.js` 容易造成 `office_strings.js`、`MicrosoftAjax.js` 等相依檔缺失。
- 檔案仍透過本地靜態伺服器提供任務窗格、manifest 與圖示資源；頁面初始化時會先驗證 `marked` 與 `turndown` 是否成功載入，若載入失敗會即時回報。

## 支援環境

目前專案分成兩條使用路徑：

- `Windows 本機版`
  - 適用：Windows + Word Desktop + sideload 測試
  - 特色：可用 launcher bridge 把 `.md` 交接給 add-in
  - 依賴：`http://localhost:3000`、Windows registry、Office sideload
- `線上可上架版`
  - 適用：正式 HTTPS 網域上的 Office Add-in 部署與商店提交流程
  - 特色：保留 taskpane 內手動匯入/匯出，不包含 Windows `.md` 雙擊關聯
  - 輸出：`dist/manifest.store.xml`

建議直接看這三份文件：

- [docs/single-machine.md](docs/single-machine.md)：Windows 單機版
- [docs/publish-online.md](docs/publish-online.md)：線上可上架版
- [docs/release-checklist.md](docs/release-checklist.md)：發布前檢查清單

## Windows `.md` 檔案關聯（可選）

以下腳本可在測試機器上建立使用者層級 `.md` 關聯。關聯不再直接把 `.md` 當 Word 文件打開，而是交給本機 launcher：launcher 會把 Markdown 內容暫存到專案內 `.local/pending-open.json`、必要時啟動本機 dev server，再開啟空白 Word 視窗交給 add-in 匯入。

```powershell
# 1) 將 .md 關聯到 Word
powershell -ExecutionPolicy Bypass -File .\scripts\enable-md-association.ps1

# 2) 還原關聯
powershell -ExecutionPolicy Bypass -File .\scripts\enable-md-association.ps1 -Undo
```

注意：這是系統層關聯，仍需 Word add-in 已 sideload。第一次使用時，若 taskpane 尚未自動顯示，請在 Word 中手動打開 add-in；之後它會讀取 launcher 交接過來的 Markdown。

## 限制與後續加強

- Manifest 沒有真正提供 Word 原生雙擊 `.md` 外掛建立/開啟流；目前改成「launcher 交接內容，再由 add-in 匯入」模式。
- 之後可加上：
  - `.md` 拖放支援
  - 直接下載 `.md` 檔（包含 BOM 與檔名）
  - 更完整段落格式還原（標題、清單、表格、程式碼區塊）
  - 單元測試與自動化驗證（Playwright + Office 驗證腳本）
