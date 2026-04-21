# Word Markdown Companion Add-in

這個 repo 目前同時維護兩種模式：

- `單機版`
  - 給 Windows + Word Desktop + sideload 測試
  - 包含 `.md` 關聯與 launcher bridge
- `線上版`
  - 給正式 HTTPS 網域上的 Office Add-in 部署或上架流程
  - 不包含 Windows shell 關聯

## 入口文件

- [docs/single-machine.md](docs/single-machine.md)
- [docs/publish-online.md](docs/publish-online.md)
- [docs/online-smoke-test.md](docs/online-smoke-test.md)
- [docs/github-pages.md](docs/github-pages.md)
- [docs/release-checklist.md](docs/release-checklist.md)
- [docs/skill-list.md](docs/skill-list.md)

## 一步完成指令

### 單機版

```bash
npm run single-machine
```

### 線上版

先設定正式 HTTPS host：

```powershell
$env:MANIFEST_HOST = "https://your-addin-host.example"
$env:SUPPORT_URL = "https://your-addin-host.example/support"
```

再執行：

```bash
npm run online
```

這會輸出：

- `dist/manifest.store.xml`
- `dist/site/`

### GitHub Pages

這個 repo 目前採 `gh-pages` branch 發佈，不依賴 GitHub Actions。發佈時：

- 先設定 GitHub Pages 實際網址為 `MANIFEST_HOST`
- 執行 `npm run build:online`
- 把 `dist/site/` 內容發佈到 `gh-pages` branch root
- Pages 站點會提供：
  - `index.html`
  - `manifest.store.xml`
  - `taskpane.html`
  - `js/*`, `styles/*`, `lib/*`, `locales/*`, `assets/*`

## 重要說明

- `單機版` 依賴 `localhost`、Windows registry 與 Word Desktop sideload，會從 `3000` 起自動選可用 port。
- `線上版` 會輸出正式版 manifest 與可直接部署到靜態主機的 `dist/site/`，不包含本機 shell integration。
- 已停用但保留的舊流程檔案會放在 `.clean/legacy/`。

## 驗證

```bash
npm test
```

- 驗證 manifest 生成
- 驗證線上版靜態網站工件輸出
- 驗證本機 dev server 的靜態資源與 API
- 驗證 release checklist 內的核心 `node --check` 檔案
