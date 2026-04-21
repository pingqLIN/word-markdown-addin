# GitHub Pages 發佈說明

這份文件描述如何把本專案做成 `全公開 + GitHub Pages`。

## 目標

公開後預設會有兩個主要入口：

- Repo：
  - `https://github.com/pingqLIN/word-markdown-addin`
- GitHub Pages：
  - `https://pingqLIN.github.io/word-markdown-addin`

其中 Pages 站點會提供：

- `index.html`
- `manifest.store.xml`
- `taskpane.html`
- `js/*`
- `styles/*`
- `lib/*`
- `locales/*`
- `assets/*`

## 自動部署方式

本 repo 使用：

- `.github/workflows/deploy-pages.yml`

流程是：

1. push 到 `main`
2. GitHub Actions 執行 `npm ci`
3. 以 GitHub Pages URL 當作 `MANIFEST_HOST`
4. 執行 `npm run build:online`
5. 將 `dist/site/` 發佈到 GitHub Pages

## Pages 路徑

這個 repo 屬於 `project site`，不是 `user site`，所以預設網址會帶 repo 名稱：

```text
https://pingqLIN.github.io/word-markdown-addin
```

因此 manifest 內的 task pane URL 也會是：

```text
https://pingqLIN.github.io/word-markdown-addin/taskpane.html
```

## 公開前檢查

1. 確認 repo 內容可以公開。
2. 確認 `origin/main..main` 的未推送 commits 也可以公開。
3. 確認沒有機密資訊、測試帳號、私人文件殘留。
4. 確認 GitHub Actions 在此 repo 可執行。

## 公開後檢查

1. Repo visibility 變成 `Public`
2. Actions workflow 成功
3. Pages 站點可開：
   - `/`
   - `/manifest.store.xml`
   - `/taskpane.html`
4. `manifest.store.xml` 內的網址都指向 GitHub Pages
5. Word 可以用 Pages 上的 manifest 載入 add-in
