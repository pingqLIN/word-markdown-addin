# GitHub Pages 發佈說明

這份文件描述目前這個 repo 如何用 `全公開 + GitHub Pages` 提供 Office Add-in 靜態資源。

## 目前採用的方式

本 repo 目前採 `branch publish`，不依賴 GitHub Actions workflow。

原因：

- Pages 內容本質上就是 `dist/site/` 純靜態檔
- 目前帳號的 Actions 執行會被 billing lock 擋住
- `gh-pages` branch root 發佈已足夠支撐這個 add-in 的公開 host

## 目前正式入口

- Repo:
  - `https://github.com/pingqLIN/word-markdown-addin`
- GitHub Pages:
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

## Pages 站點會提供的內容

- `index.html`
- `install.html`
- `support.html`
- `privacy.html`
- `manifest.store.xml`
- `taskpane.html`
- `js/*`
- `styles/*`
- `lib/*`
- `locales/*`
- `assets/*`

## 發佈流程

1. 先確認 Pages 設定：

```powershell
gh api repos/pingqLIN/word-markdown-addin/pages
```

預期至少要看到：

- `html_url = https://github.colorgeek.co/word-markdown-addin/`
- `source.branch = gh-pages`
- `source.path = /`
- `build_type = legacy`

2. 設定建置用公開網址：

```powershell
$env:MANIFEST_HOST = "https://github.colorgeek.co/word-markdown-addin"
$env:SUPPORT_URL = "https://github.colorgeek.co/word-markdown-addin/support.html"
```

若未來取得 Marketplace asset ID，可再加：

```powershell
$env:MARKETPLACE_ASSET_ID = "WA000000000"
$env:MARKETPLACE_ADDIN_TITLE = "Word Markdown Companion"
$env:MARKETPLACE_LINK_LANGUAGE = "en-US"
```

3. 產生公開站點工件：

```powershell
npm run build:online
```

4. 把 `dist/site/` 全部內容發到 `gh-pages` branch root。

## Pages 路徑規則

- `manifest.store.xml` 必須指回同一個 Pages host
- task pane URL 必須是：

```text
https://github.colorgeek.co/word-markdown-addin/taskpane.html
```

- install page URL 會是：

```text
https://github.colorgeek.co/word-markdown-addin/install.html
```

- support page URL 會是：

```text
https://github.colorgeek.co/word-markdown-addin/support.html
```

也就是說，使用者應該先到 `install.html`，Word 再透過 manifest 去載 `taskpane.html`。

## 公開前檢查

1. 確認 repo 內容可以公開。
2. 確認 `origin/main..main` 的未推送 commits 也可以公開。
3. 確認沒有機密資訊、測試帳號、私人文件殘留。
4. 確認 `gh-pages` 發佈內容是 `dist/site/`，不是 repo 原始碼。

## 公開後檢查

1. Repo visibility 是 `Public`
2. Pages source 是 `gh-pages /`
3. Pages 站點可開：
   - `/`
   - `/install.html`
   - `/support.html`
   - `/privacy.html`
   - `/manifest.store.xml`
   - `/taskpane.html`
4. `manifest.store.xml` 內的網址都指向 GitHub Pages
5. Word 可以用 Pages 上的 manifest 載入 add-in
