# 線上版安裝說明

這份文件只整理目前公開版的安裝入口與實際網址。

## 目前正式網址

- Public site:
  - `https://github.colorgeek.co/word-markdown-addin/`
- Install page:
  - `https://github.colorgeek.co/word-markdown-addin/install.html`
- Manifest:
  - `https://github.colorgeek.co/word-markdown-addin/manifest.store.xml`
- Task pane:
  - `https://github.colorgeek.co/word-markdown-addin/taskpane.html`
- Support:
  - `https://github.colorgeek.co/word-markdown-addin/support.html`
- Privacy:
  - `https://github.colorgeek.co/word-markdown-addin/privacy.html`
- GitHub repo:
  - `https://github.com/pingqLIN/word-markdown-addin`

## 現在可用的安裝方式

### 個人安裝

1. 打開 install page。
2. 下載 `manifest.store.xml`，或直接複製 manifest URL。
3. 在 Word 內開啟 `My Add-ins`。
4. 使用 `Upload My Add-in` 載入 manifest。

這是目前對一般使用者最快可用的公開安裝方式。

### Tenant Admin 安裝

1. 打開 install page。
2. 複製部署用 manifest URL。
3. 在 Microsoft 365 admin center 以官方 manifest URL 路徑部署。

這條路徑適合內部團隊或組織級分發。

## install.html 目前做的事

`install.html` 現在會提供：

- manifest 下載按鈕
- manifest URL 複製按鈕
- admin 部署用 URL 複製按鈕
- support / privacy 頁面入口
- 官方 Microsoft 安裝說明連結
- 官方 admin deployment 文件連結

也就是說，現在它已經是「公開版安裝入口」，只是還不是 Microsoft Marketplace 意義上的真正 click-and-run。

## 真正的一鍵安裝何時成立

只有在 add-in 取得 Microsoft Marketplace asset ID 後，`install.html` 才能升級成官方 click-and-run 頁面，包含：

- Word on the web 安裝連結
- Word Desktop `ms-word:` 安裝連結

建置時只要補上以下環境變數即可：

```powershell
$env:MARKETPLACE_ASSET_ID = "WA000000000"
$env:MARKETPLACE_ADDIN_TITLE = "Word Markdown Companion"
$env:MARKETPLACE_LINK_LANGUAGE = "en-US"
```

## 公開版建置

```powershell
cd Q:\Projects\word-markdown-addin
$env:MANIFEST_HOST = "https://github.colorgeek.co/word-markdown-addin"
$env:SUPPORT_URL = "https://github.colorgeek.co/word-markdown-addin/support.html"
npm run build:online
```

產生的安裝相關工件：

- `dist/manifest.store.xml`
- `dist/site/index.html`
- `dist/site/install.html`
- `dist/site/support.html`
- `dist/site/privacy.html`
- `dist/site/taskpane.html`

## 重要區分

- `install.html` 是公開版安裝入口
- `taskpane.html` 是 Office Add-in 實際載入的 UI 入口
- 一般瀏覽器使用者不應該把 `taskpane.html` 當成完整獨立網站使用
