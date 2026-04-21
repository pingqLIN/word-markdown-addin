# 線上版 Smoke Test 指南

這份文件用於 `官方 / 線上版` 路徑的實機驗證。

目標不是測 `.md` 副檔名關聯，而是驗證：

- 正式 HTTPS host 上的 taskpane 是否可載入
- Word Desktop / Word Online 是否能正常使用 add-in
- 匯入、匯出、語系切換、基本互動是否正常
- `Office.AutoShowTaskpaneWithDocument` 相關行為是否至少在可支援情境下成立

## 前置條件

1. 已設定正式 HTTPS host：

```powershell
$env:MANIFEST_HOST = "https://your-addin-host.example"
$env:SUPPORT_URL = "https://your-addin-host.example/support"
```

2. 已執行：

```bash
npm run online
```

3. 已將 `dist/site/` 內容部署到正式 HTTPS host 根目錄。
4. `dist/manifest.store.xml` 已指向正確的正式 host。
5. 準備測試樣本檔案：
   - [samples/official-smoke-sample.md](Q:\Projects\word-markdown-addin\samples\official-smoke-sample.md)

## Smoke Test 範圍

這份 smoke test 只驗證官方主路徑：

- Word Desktop
- Word Online
- taskpane UI
- Markdown 匯入 / 匯出

這份 smoke test 不驗證：

- Windows `.md` 關聯
- `scripts/enable-md-association.ps1`
- `scripts/open-markdown-in-word.js`
- localhost bridge

## Word Desktop

### A. 載入驗證

1. 關閉所有 Word 視窗。
2. 用 `dist/manifest.store.xml` 載入 add-in。
3. 重新開啟 Word。
4. 確認 ribbon 上能看到 `Markdown` / `Markdown Tools` 群組。

預期結果：

- add-in 可被載入
- taskpane 可被打開
- taskpane 內容來自正式 HTTPS host，不是 localhost

### B. 匯入驗證

1. 在 Word Desktop 開啟空白文件。
2. 打開 add-in taskpane。
3. 使用 `Import .md` 或 taskpane 內的匯入按鈕。
4. 選取 [official-smoke-sample.md](Q:\Projects\word-markdown-addin\samples\official-smoke-sample.md)。

預期結果：

- 內容可插入目前文件
- 標題、清單、表格、強調、連結至少能以合理 Word 格式呈現
- taskpane 狀態訊息不出現 `Office.context` / `Office.js` / hosted asset 失敗

### C. 匯出驗證

1. 在同一份文件中點 `Export .md` 或 taskpane 內的匯出操作。
2. 檢查預覽區內容。
3. 測試：
   - `選擇位置另存 MD 檔`
   - `一鍵複製全文`

預期結果：

- 預覽區出現 Markdown
- 下載或另存可成功
- 複製按鈕可成功寫入剪貼簿

### D. 自動顯示 taskpane 驗證

1. 匯入完成後，把文件存成 `.docx`。
2. 關閉文件。
3. 重新開啟剛才存下的 `.docx`。

預期結果：

- 若 Word / 文件狀態支援，taskpane 會自動顯示或可由同一 add-in 快速回到上次工作流
- 若未自動顯示，但 add-in 本身仍可正常從 ribbon 重新開啟，記錄為 `best-effort fallback`，不要直接判定整體失敗

## Word Online

### A. 載入驗證

1. 在瀏覽器開啟 Word Online。
2. 以正式發佈或內部分發方式讓 add-in 可被使用。
3. 開啟空白文件或既有 `.docx`。
4. 打開 add-in taskpane。

預期結果：

- taskpane 可正常載入
- 靜態資源都來自正式 HTTPS host
- 沒有依賴 Windows registry 或 localhost

### B. 匯入 / 匯出驗證

1. 匯入 [official-smoke-sample.md](Q:\Projects\word-markdown-addin\samples\official-smoke-sample.md)。
2. 檢查文件內呈現。
3. 執行匯出。
4. 檢查預覽、下載或複製功能。

預期結果：

- 匯入與匯出核心路徑可運作
- 語系切換、基本 UI 操作正常

## 失敗分流

### taskpane 打不開

先分辨：

- manifest 問題
- 正式 HTTPS host 資源缺檔
- Office 端 add-in 載入問題

先檢查：

- `dist/manifest.store.xml`
- `dist/site/` 是否完整部署
- `/taskpane.html`
- `/js/taskpane.js`
- `/styles/taskpane.css`
- `/lib/marked.min.js`
- `/lib/turndown.min.js`
- `/locales/zh-TW.json`
- `/locales/en-US.json`
- `/assets/icon-16.png`
- `/assets/icon-32.png`
- `/assets/icon-80.png`

### 匯入失敗

先分辨：

- Office.js / Word API 問題
- Markdown 樣本內容問題
- taskpane UI 事件沒有觸發

### 匯出失敗

先分辨：

- 文件內容為空
- Word `getHtml()` 回傳不符合預期
- 瀏覽器下載 / clipboard 權限限制

## 測試紀錄建議

每次 smoke test 至少記：

- 測試日期
- 測試環境
  - Word Desktop 或 Word Online
  - host URL
- manifest 版本
- 通過項目
- 失敗項目
- 是否屬於 `best-effort` 行為差異

## 完成判定

以下都成立時，可視為線上版 smoke test 基本通過：

- Word Desktop 可載入 add-in
- Word Online 可載入 add-in
- 匯入 `.md` 正常
- 匯出 `.md` 正常
- 正式 host 資源完整
- 不依賴 registry、localhost 或 launcher bridge
