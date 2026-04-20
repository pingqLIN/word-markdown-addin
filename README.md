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
- [docs/release-checklist.md](docs/release-checklist.md)

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

## 重要說明

- `單機版` 依賴 `localhost`、Windows registry 與 Word Desktop sideload。
- `線上版` 只輸出正式版 manifest 到 `dist/manifest.store.xml`，不包含本機 shell integration。
- 已停用但保留的舊流程檔案會放在 `.clean/legacy/`。
