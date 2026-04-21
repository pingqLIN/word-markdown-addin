# 線上版使用說明

這份文件只描述 `線上可上架版`。

## 適用環境

- 正式 HTTPS 網域上的 Office Add-in 部署
- Word Desktop 與 Word Online 的正式載入流程
- AppSource 或內部 catalog 提交流程
- Node.js 20+

## 線上版保留的能力

- 任務窗格 UI
- 手動選取 `.md` 檔匯入
- 拖放 `.md` 檔匯入
- 將目前文件匯出成 Markdown 並下載

## 線上版不包含

- Windows `.md` 檔關聯
- `scripts/enable-md-association.ps1`
- `scripts/open-markdown-in-word.js`
- `.local/pending-open.json` bridge
- 任何依賴 `localhost` 或 registry 的流程

## 一步完成安裝與發布包準備

先設定正式 HTTPS host：

```powershell
$env:MANIFEST_HOST = "https://your-addin-host.example"
$env:SUPPORT_URL = "https://your-addin-host.example/support"
```

再在 repo 根目錄執行：

```bash
npm run online
```

這個指令會依序完成：

1. 安裝 npm 相依
2. 產生正式 HTTPS 版 manifest
3. 輸出到 `dist/manifest.store.xml`
4. 建立可部署到靜態主機的 `dist/site/`

## 輸出檔案

- `dist/manifest.store.xml`
- `dist/site/taskpane.html`
- `dist/site/js/*`
- `dist/site/styles/*`
- `dist/site/lib/*`
- `dist/site/locales/*`
- `dist/site/assets/*`

建議搭配測試樣本：

- [samples/official-smoke-sample.md](Q:\Projects\word-markdown-addin\samples\official-smoke-sample.md)

若要覆寫 metadata，也可以先設定：

```powershell
$env:ADDIN_ID = "00000000-0000-0000-0000-000000000000"
$env:PROVIDER_NAME = "Your Company"
$env:DISPLAY_NAME = "Word Markdown Companion"
$env:ADDIN_DESCRIPTION = "Import and export Markdown files in Microsoft Word."
```

## 上架前檢查重點

1. `MANIFEST_HOST` 必須是正式 HTTPS 網域。
2. `SUPPORT_URL` 必須是正式 HTTPS 網址。
3. 正式 host 必須能提供：
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
4. 不要把 Windows shell 關聯描述成線上版功能。
5. 用正式 host 重新產生 manifest 後，再做 Word Desktop 與 Word Online smoke test。

## 部署建議

1. 先把 `dist/site/` 內容部署到你的正式 HTTPS host 根目錄。
2. 再使用 `dist/manifest.store.xml` 做 Word Desktop / Word Online 載入與後續上架或內部分發。
3. 若主機不是部署在根目錄，請先確保對外路徑仍能對應到 manifest 內的 `/taskpane.html`、`/js/*`、`/styles/*`、`/lib/*`、`/locales/*`、`/assets/*`。

更完整的實機驗證步驟請看：

- [docs/online-smoke-test.md](Q:\Projects\word-markdown-addin\docs\online-smoke-test.md)
- [docs/github-pages.md](Q:\Projects\word-markdown-addin\docs\github-pages.md)
