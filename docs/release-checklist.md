# 發布前檢查清單

這份清單用來確認 repo 是否已整理到可提交或可對外發布的狀態。

## 1. Repo 狀態

- `git status --short --branch` 確認 worktree 已整理完成
- 不要把不必要的 `.local/` 測試殘留、臨時輸出或診斷檔一併帶入發布
- 確認 `dist/manifest.store.xml` 是用正式 HTTPS host 重新生成

## 2. 文件一致性

- README 與 docs 是否已說清楚：
  - `Windows 單機版`
  - `線上可上架版`
- 不要把 Windows 專用 launcher 功能寫成上架版承諾
- icon、manifest、腳本名稱與實際檔案一致

## 3. Manifest 與資源

- `manifest.xml` 用於本機 sideload
- `dist/manifest.store.xml` 用於線上版或上架流程
- `dist/site/` 用於正式 HTTPS host 的靜態網站部署
- 正式版 manifest 必須全部指向 HTTPS
- 正式版 manifest 的 `SupportUrl` 應指向公開可用的 `support.html`
- icon、taskpane、js、css、lib 路徑都可由正式 host 提供
- 若走 GitHub Pages，確認 `dist/site/` 內包含 `index.html`、`install.html`、`support.html`、`privacy.html`、`.nojekyll`、`manifest.store.xml`

## 4. 核心驗證

- `npm test`
- `npm run build:site`
- `gh api repos/pingqLIN/word-markdown-addin/pages`
- `node --check scripts/render-manifest.js`
- `node --check scripts/build-static-site.js`
- `node --check scripts/dev-server.js`
- `node --check scripts/project-loop-state.js`
- `node --check scripts/setup-auto.js`
- `node --check scripts/sideload.js`
- `node --check src/js/taskpane.js`

## 5. 單機版 smoke test

- `npm run single-machine`
- 參考 `samples/official-smoke-sample.md`
- 從 Word 開啟 taskpane
- 手動選檔匯入 `.md`
- 拖放 `.md` 匯入
- 匯出 Markdown 並下載
- 若有啟用關聯，雙擊 `.md` 後確認 bridge 匯入成功

## 6. 線上版 smoke test

- 設定正式 `MANIFEST_HOST`
- 設定正式 `SUPPORT_URL` 指向 `support.html`
- 執行 `npm run online`
- 部署 `dist/site/` 到正式 HTTPS host
- 參考 `docs/online-smoke-test.md`
- 使用 `samples/official-smoke-sample.md`
- 用正式 HTTPS manifest 載入 add-in
- 測試 Word Desktop
- 測試 Word Online
- 確認不依賴 registry、`.local` 或 launcher bridge

## 7. 發布判定

只有在以下條件都成立時，才算接近可發布：

- worktree 已整理
- 文件與功能一致
- 正式版 manifest 已生成
- `support.html` / `privacy.html` 已跟公開 host 一起部署
- Windows 單機版 smoke test 通過
- 線上版 smoke test 通過
- 已知限制已寫清楚
