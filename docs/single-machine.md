# 單機版使用說明

這份文件只描述 `Windows 單機版`。

## 適用環境

- Windows 10/11
- Microsoft Word Desktop
- Office add-in sideload 測試環境
- Node.js 20+

## 單機版保留的能力

- 在 taskpane 內手動選取 `.md` 檔匯入
- 拖放 `.md` 檔匯入
- 將目前 Word 文件匯出成 Markdown
- 用 Windows shell 關聯與 launcher bridge，讓雙擊 `.md` 時把內容交接給 add-in

## 一步完成安裝與啟動

在 repo 根目錄執行：

```bash
npm run single-machine
```

這個指令會依序完成：

1. 安裝 npm 相依
2. 產生本機版 `manifest.xml`
3. 啟用 `.md` 關聯與 launcher bridge
4. 從 `3000` 起自動選一個可用 localhost port，並啟動或重用對應的 dev server
5. 將 add-in sideload 到 Word Desktop

執行完成後：

- 關掉所有 Word 視窗再重開
- Word 會把匯入 taskpane 設成預設 pane，通常在安裝後或文件再次開啟時自動顯示
- 若 ribbon 沒立即出現 add-in，先到 Word 的 Add-ins 手動開一次

## 單機版依賴

- `http://localhost:<auto-selected-port>`
- Windows registry
- Office sideload registry
- `.local/pending-open.json` bridge
- `.local/runtime.json`

## 重要限制

- 這不是 Word 原生支援 Markdown 開檔。
- `.md` 不能直接被當成 Word 文件打開；必須先經過 launcher bridge。
- 某些 Word / WebView 環境仍可能對 `http://localhost` 有限制。
- 自動開啟 taskpane 屬於 Office 端 best-effort 行為，不同安裝管道與文件狀態可能略有差異。
- 若桌面版 Word 出現 `We can't open this add-in from localhost`，仍需替 WebView 開 loopback exemption。

## 適合用途

- 本機功能開發
- Windows Word Desktop 測試
- 內部原型驗證
- launcher bridge 流程驗證
