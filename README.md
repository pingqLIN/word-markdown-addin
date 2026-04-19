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
- `assets/icon.svg`：add-in 圖示。
- `scripts/dev-server.js`：本地靜態伺服器。
- `src/lib/office.js`、`src/lib/marked.min.js`、`src/lib/turndown.min.js`：本地核心腳本，避免啟動時依賴外部 CDN。
- `package.json`：專案指令。

## 快速開始

1. 安裝 Node.js 20+。
2. 在根目錄執行：
   ```bash
   npm install
   npm run dev
   ```
   服務會在 `http://localhost:3000` 啟動（建議本機先用 HTTP 測試；上線或企業環境可改 HTTPS）。
3. 在 Word 選擇「上傳我的清單」/sideload manifest，載入 `manifest.xml`。

## 使用方式

1. 在 Word 開啟文件後，從 add-in 按鈕開啟任務窗格。
2. 點「插入 Markdown 到文件」：
   - 選擇 `.md` 檔案。
   - 檔案內容會被轉為 HTML 插入到游標位置。
3. 點「匯出為 Markdown」：
   - 點選後會把目前 Word 文件轉為 Markdown 字串顯示在文字區。
   - 點「下載為 Markdown 檔」可直接儲存成 `.md`。

## 注意事項

- 這是 Word 任務窗格層級的「導入／導出」整合，能讓使用者在 Word 內直接操作 `.md`。  
- 若要做到「作業系統原生」雙擊 `.md` 就直接在 Word 透過這個增益集開啟，屬於「檔案關聯/IT 部署政策」層級議題，不是 manifest 本身可以直接保證的行為。

## 目前狀態更新

- 本地化流程：`office.js`、`marked` 與 `turndown` 都改為以 `src/lib` 載入，降低啟動與轉換階段對外部 CDN 的依賴。
- 檔案仍透過本地靜態伺服器提供任務窗格、圖示與腳本資源；頁面初始化時會先驗證 `marked` 與 `turndown` 是否成功載入，若載入失敗會即時回報。

## Windows `.md` 檔案關聯（可選）

以下腳本可在測試機器上建立使用者層級 `.md` 關聯，讓 `.md` 預設用 Word 開啟（Word 會接手後再用本增益集匯入）。

```powershell
# 1) 將 .md 關聯到 Word
powershell -ExecutionPolicy Bypass -File .\scripts\enable-md-association.ps1

# 2) 還原關聯
powershell -ExecutionPolicy Bypass -File .\scripts\enable-md-association.ps1 -Undo
```

注意：這是系統層關聯，不會在不啟動 Office Add-in 的情況下自動將 .md 內容直接注入 Word；仍需加值集側載與開啟 taskpane。

## 限制與後續加強

- Manifest 沒有真正提供 Word 原生雙擊 `.md` 外掛建立/開啟流；目前是「載入增益集後，由使用者手動匯入」模式。
- 之後可加上：
  - `.md` 拖放支援
  - 直接下載 `.md` 檔（包含 BOM 與檔名）
  - 更完整段落格式還原（標題、清單、表格、程式碼區塊）
  - 單元測試與自動化驗證（Playwright + Office 驗證腳本）
