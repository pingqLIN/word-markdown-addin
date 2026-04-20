# 單機版使用說明

這份文件描述目前 repo 的 `Windows 單機版` 使用方式。

## 適用環境

- Windows 10/11
- Microsoft Word Desktop
- Office add-in sideload 開發/測試環境
- Node.js 20+

## 核心能力

- 在 taskpane 內手動選取 `.md` 檔匯入
- 拖放 `.md` 檔匯入
- 將目前 Word 文件匯出成 Markdown
- 使用 Windows shell 關聯與 launcher bridge，讓雙擊 `.md` 時把內容交接給 add-in

## 依賴項

- `http://localhost:3000`
- Windows registry
- Office sideload registry
- `.local/pending-open.json` bridge

## 啟動方式

```bash
npm run setup
npm run sideload
```

如果要同時啟用 `.md` 檔關聯：

```bash
npm run setup:with-md-association
```

## 重要限制

- 這不是 Word 原生支援 Markdown 開檔。
- `.md` 不能直接被當成 Word 文件打開；必須先經過 launcher bridge。
- 某些 Word / WebView 環境仍可能對 `http://localhost` 有限制。

## 適合用途

- 本機功能開發
- Windows Word Desktop 測試
- 內部原型驗證
- launcher bridge 流程驗證
