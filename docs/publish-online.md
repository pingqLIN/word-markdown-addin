# 線上可上架版本

這個專案現在分成兩種使用模式：

- `Windows 本機版`
  - 給 Windows + Word Desktop + sideload 測試
  - 依賴 `http://localhost:3000`
  - 可選擇 `.md` 檔案關聯與 launcher bridge
- `線上可上架版`
  - 給 HTTPS 網域上的 Office Add-in 部署或 AppSource 提交流程
  - 不依賴 Windows registry、`localhost` 或 `.md` 雙擊關聯
  - 主要能力是 taskpane 內的手動匯入 `.md` 與匯出 `.md`

## 線上版的功能邊界

線上可上架版本保留：

- 任務窗格 UI
- 手動選取 `.md` 檔匯入
- 拖放 `.md` 檔匯入
- 將目前文件匯出成 Markdown 並下載

線上可上架版本不包含：

- `scripts/enable-md-association.ps1`
- `scripts/open-markdown-in-word.js`
- `.local/pending-open.json` bridge
- 任何需要本機 registry 或 Word Desktop shell integration 的流程

這些功能是 Windows 本機版專用，不應列入上架版承諾。

## 產生線上版 manifest

先在 PowerShell 設定你的正式 HTTPS host：

```powershell
$env:MANIFEST_HOST = "https://your-addin-host.example"
$env:SUPPORT_URL = "https://your-addin-host.example/support"
npm run build:store
```

輸出檔案：

- `dist/manifest.store.xml`

若你要覆寫 metadata，也可以另外設定：

```powershell
$env:ADDIN_ID = "00000000-0000-0000-0000-000000000000"
$env:PROVIDER_NAME = "Your Company"
$env:DISPLAY_NAME = "Word Markdown Companion"
$env:ADDIN_DESCRIPTION = "Import and export Markdown files in Microsoft Word."
```

## 上架前檢查

1. `MANIFEST_HOST` 必須是正式 HTTPS 網域。
2. `SUPPORT_URL` 必須是正式 HTTPS 網址。
3. 該 HTTPS 網域必須能提供：
   - `/taskpane.html`
   - `/js/taskpane.js`
   - `/styles/taskpane.css`
   - `/lib/marked.min.js`
   - `/lib/turndown.min.js`
   - `/assets/icon-16.png`
   - `/assets/icon-32.png`
   - `/assets/icon-80.png`
4. 不要把 README 中的 Windows shell 關聯描述當成上架版功能。
5. 重新用正式 host 產生 manifest 後，再做一次 Word Desktop 與 Word Online smoke test。

## 建議 smoke test

1. 在 Word Desktop 以正式 HTTPS manifest 載入 add-in。
2. 在 Word Online 載入同一份 manifest。
3. 測試手動選檔匯入 `.md`。
4. 測試拖放 `.md` 匯入。
5. 測試匯出並下載 `.md`。
6. 確認沒有任何 `localhost`、registry、或 `.local` bridge 依賴。
