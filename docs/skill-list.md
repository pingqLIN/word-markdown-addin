# Skill List

這份清單整理目前這個 workspace 中，對 `word-markdown-addin` 最有實際價值的 skills。

它不是整個執行環境的完整技能總表，而是 repo-local 的工作清單，用來幫助後續維護時快速判斷應該叫用哪一類能力。

## 目前建議使用的 skills

### UI / 前端

- `frontend-design`
  - 用於 taskpane 版面重配、按鈕強化、資訊層級重組、互動區塊視覺整理。
  - 這個 skill 是目前最接近 `$claude-design-playbook` 的替代方案。

- `webapp-testing`
  - 用於本機 taskpane 頁面或其他 localhost 頁面的前端互動驗證。
  - 適合檢查切換、顯示、拖放、輸出預覽等 UI 流程。

### 文件 / 工作流

- `doc-coauthoring`
  - 用於整理專案說明、部署說明、release checklist、維護文件。

- `internal-comms`
  - 用於撰寫狀態更新、變更摘要、交接說明、內部公告式文件。

### 環境 / 能力盤點

- `env`
  - 用於查看目前載入的 skills、apps、plugins、MCP servers 與其他執行環境資訊。
  - 當需要確認「這個 skill 現在有沒有真的可用」時，優先用它。

### 測試 / 品質

- `project-development-loop`
  - 適合做發布前整理、進度盤點、風險檢查、收尾 review。
  - 如果要做 release 前 audit，優先用這個。

## 對本專案目前最相關的 skill mapping

- `taskpane UI redesign`
  - 優先使用 `frontend-design`

- `taskpane smoke test`
  - 優先使用 `webapp-testing`

- `README / docs / release notes 更新`
  - 優先使用 `doc-coauthoring`

- `發布前整理檢查`
  - 優先使用 `project-development-loop`

- `確認目前 skill / plugin / app 是否存在`
  - 優先使用 `env`

## 目前不可用或未安裝的項目

- `$claude-design-playbook`
  - 目前不在這個 workspace 的可用 skill 清單中。
  - 若使用者提到這個名稱，先回退到 `frontend-design`，除非之後真的安裝進來。

## 維護建議

- 如果未來新增 repo 專屬 skill，請把用途與觸發時機補在這份文件。
- 如果某個全域 skill 常被這個 repo 使用，也可以補進這裡。
- 不要把整個全域 skill registry 原封不動複製進 repo；只保留真正有助於本專案維護的子集。
