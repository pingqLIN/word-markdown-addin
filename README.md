# Word Markdown Companion

Public Microsoft Word add-in for importing and exporting Markdown `.md` files inside Word.

This repo keeps two tracks in parallel:

- `線上版`
  - 正式公開路徑
  - 依賴 HTTPS host、manifest 與 Office Add-in 標準分發
- `單機版`
  - Windows + Word Desktop 本機 helper 路徑
  - 給開發、測試與 sideload 使用，不是公開發佈主路徑

## 已上線網址

- Public site: `https://github.colorgeek.co/word-markdown-addin/`
- Install page: `https://github.colorgeek.co/word-markdown-addin/install.html`
- Manifest: `https://github.colorgeek.co/word-markdown-addin/manifest.store.xml`
- Task pane: `https://github.colorgeek.co/word-markdown-addin/taskpane.html`
- Support: `https://github.colorgeek.co/word-markdown-addin/support.html`
- Privacy: `https://github.colorgeek.co/word-markdown-addin/privacy.html`
- GitHub repo: `https://github.com/pingqLIN/word-markdown-addin`

## Mascot

![Word Markdown Companion mascot](assets/mascot/word-markdown-companion-url-hero.jpg)

- 生成來源：[assets/mascot/word-markdown-companion-url-hero.md](assets/mascot/word-markdown-companion-url-hero.md)
- 圖檔：[assets/mascot/word-markdown-companion-url-hero.jpg](assets/mascot/word-markdown-companion-url-hero.jpg)

## 線上版安裝

目前公開版請從 install page 開始：

- 個人安裝：下載 manifest，然後在 Word 的 `My Add-ins` / `Upload My Add-in` 載入
- Tenant Admin 安裝：在 Microsoft 365 admin center 使用 manifest URL 部署

真正的 Microsoft 官方 click-and-run 一鍵安裝，必須等這個 add-in 取得 Marketplace asset ID 後才能啟用。現在的 `install.html` 已經把這個入口預留好了。

## 建置公開版

```powershell
cd Q:\Projects\word-markdown-addin
$env:MANIFEST_HOST = "https://github.colorgeek.co/word-markdown-addin"
$env:SUPPORT_URL = "https://github.colorgeek.co/word-markdown-addin/support.html"
npm run build:online
```

輸出內容：

- `dist/manifest.store.xml`
- `dist/site/index.html`
- `dist/site/install.html`
- `dist/site/support.html`
- `dist/site/privacy.html`
- `dist/site/taskpane.html`
- `dist/site/js/*`
- `dist/site/styles/*`
- `dist/site/lib/*`
- `dist/site/locales/*`
- `dist/site/assets/*`

## 文件入口

### 公開版 / 線上版

- [docs/online-install.md](docs/online-install.md)
- [docs/publish-online.md](docs/publish-online.md)
- [docs/github-pages.md](docs/github-pages.md)
- [docs/online-smoke-test.md](docs/online-smoke-test.md)

### 本機版

- [docs/single-machine.md](docs/single-machine.md)

### 維護 / 測試

- [docs/release-checklist.md](docs/release-checklist.md)
- [docs/skill-list.md](docs/skill-list.md)

## 補充

- GitHub Pages 目前是公開靜態 host，不是獨立 Web App。
- 線上版不包含 Windows `.md` 關聯、registry 或 `localhost` bridge。
- 線上版 manifest 應指向公開的 `support.html`，不要再回退到 repo root 或本機 helper 路徑。
- 若之後取得 Marketplace asset ID，只要重新建置 `install.html`，就能升級成真正的官方一鍵安裝頁。
