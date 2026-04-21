# GitHub Pages 發佈說明

這份文件描述如何把本專案做成 `全公開 + GitHub Pages`。

## 目前採用的方式

本 repo 目前採 `branch publish`，不依賴 GitHub Actions workflow。

原因很單純：

- 這個 repo 的 Pages 內容本質上就是 `dist/site/` 純靜態檔
- GitHub 官方也建議在不需要自訂 build flow 時，直接用 branch 當 publishing source
- 目前帳號的 Actions 執行會被 billing lock 擋住，workflow 發佈不可用

## 目標

公開後主要會有兩個入口：

- Repo：
  - `https://github.com/pingqLIN/word-markdown-addin`
- GitHub Pages：
  - 以 GitHub Pages API 回報的 `html_url` 為準

其中 Pages 站點會提供：

- `index.html`
- `manifest.store.xml`
- `taskpane.html`
- `js/*`
- `styles/*`
- `lib/*`
- `locales/*`
- `assets/*`

## 發佈流程

1. 先查出 Pages 實際網址：

```powershell
gh api repos/pingqLIN/word-markdown-addin/pages
```

2. 把 `html_url` 換成 HTTPS，設成 `MANIFEST_HOST`：

```powershell
$env:MANIFEST_HOST = "https://github.colorgeek.co/word-markdown-addin"
$env:SUPPORT_URL = "https://github.com/pingqLIN/word-markdown-addin"
```

3. 產生公開站點工件：

```powershell
npm run build:online
```

4. 把 `dist/site/` 全部內容發到 `gh-pages` branch root。

5. 把 GitHub Pages publishing source 設成：

- branch: `gh-pages`
- path: `/`
- build type: `legacy`

## Pages 路徑

task pane URL 會跟著 `MANIFEST_HOST`：

```text
https://github.colorgeek.co/word-markdown-addin/taskpane.html
```

manifest 也必須指向同一個 host，否則 Word 會載到錯的 task pane URL。

## 公開前檢查

1. 確認 repo 內容可以公開。
2. 確認 `origin/main..main` 的未推送 commits 也可以公開。
3. 確認沒有機密資訊、測試帳號、私人文件殘留。
4. 確認 `gh-pages` 發佈內容是 `dist/site/`，不是 repo 原始碼。

## 公開後檢查

1. Repo visibility 變成 `Public`
2. Pages source 變成 `gh-pages /`
3. Pages 站點可開：
   - `/`
   - `/manifest.store.xml`
   - `/taskpane.html`
4. `manifest.store.xml` 內的網址都指向 GitHub Pages
5. Word 可以用 Pages 上的 manifest 載入 add-in
