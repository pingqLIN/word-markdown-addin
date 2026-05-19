[![Word Markdown Companion product banner](Banner_HDR2.jpg)](https://github.colorgeek.co/word-markdown-addin/install.html)

# WORD MARKDOWN COMPANION

**Import, format, and export Markdown inside Microsoft Word**

![Version](https://img.shields.io/badge/version-0.1.0-c6543c)
![Platform](https://img.shields.io/badge/platform-Microsoft%20Word-2b579a)
![Runtime](https://img.shields.io/badge/node-%3E%3D20-43853d)
![License](https://img.shields.io/badge/license-MIT-blue)

[Live Site](https://github.colorgeek.co/word-markdown-addin/) ·
[Install](https://github.colorgeek.co/word-markdown-addin/install.html) ·
[Quick Start](#start-using) ·
[Documentation](#documentation) ·
[繁體中文](README.zh-TW.md)

---

## Why Word Markdown Companion

Word Markdown Companion brings Markdown import and export into a Word task pane, with separate paths for public HTTPS deployment and Windows desktop sideload testing. Writers can keep `.md` files in their normal workflow, then move content into Word for editing, review, or handoff.

| Benefit | What it gives you |
|---|---|
| **Markdown import** | Select or drop `.md` / `.markdown` files — insert rendered content into the current Word document |
| **Markdown export** | Convert the current Word document to Markdown — preview, copy, download, or save |
| **Public add-in path** | HTTPS manifest, support page, privacy page, static GitHub Pages host |
| **Windows local path** | Localhost dev server, sideload manifest, optional `.md` shell handoff bridge |

---

## Screenshots

| Task pane top | Task pane bottom |
|---|---|
| ![English task pane top screenshot](assets/screenshots/taskpane-en-top.png) | ![English task pane bottom screenshot](assets/screenshots/taskpane-en-bottom.png) |

The current English task pane shows the import workflow, localized file picker, drop zone, and Markdown conversion reference.

> **Brand artwork rule:** legacy mascot artwork is source material only, not a README display asset. Do not place the mascot directly in documentation or product surfaces; redesign it into approved banner, campaign, or product artwork before use.

---

## How It Works

### 1. Install or sideload the manifest

Use the public install page for the online track, or the Windows sideload command for local development.

### 2. Open the task pane in Word

The manifest points Word to `taskpane.html`, hosted publicly for the online track or served from localhost for the local track.

### 3. Import, format, or export Markdown

Use `Import .md`, drag-and-drop, `Format`, or `Export .md` from the same task pane workspace.

---

## Who It Is For

| User | What it helps you finish |
|---|---|
| Markdown-first writers | Draft in `.md`, move content into Word for review or delivery |
| Word-heavy teams | Receive Markdown source without leaving the Word editing surface |
| Add-in testers | Validate Word Desktop and Word Online behavior from a single manifest track |
| Windows local users | Test sideload, shell association, and launcher bridge flows without publishing |

---

## Support Matrix

| Type | Support |
|---|---|
| Input | `.md`, `.markdown`, common Markdown, GFM-style tables |
| Output | Markdown preview, clipboard copy, file download, save picker fallback |
| Online host | GitHub Pages static site, HTTPS manifest, public support and privacy pages |
| Word platforms | Word Desktop, Word Online, platforms allowed by Office manifest validation |
| Local development | Windows Word Desktop sideload, localhost dev server, optional `.md` association |

> **Important:** the online track does not include Windows registry edits, `.md` file association, `.local/pending-open.json`, or a localhost launcher bridge.

---

## Start Using

Open the public install page:

**[https://github.colorgeek.co/word-markdown-addin/install.html](https://github.colorgeek.co/word-markdown-addin/install.html)**

### Public HTTPS Build

```powershell
# Build the public Office Add-in bundle
cd Q:\Projects\word-markdown-addin
$env:MANIFEST_HOST = "https://github.colorgeek.co/word-markdown-addin"
$env:SUPPORT_URL = "https://github.colorgeek.co/word-markdown-addin/support.html"
npm run build:online
```

### Windows Desktop Sideload

```powershell
# Prepare the local helper path and sideload into Word Desktop
cd Q:\Projects\word-markdown-addin
npm run single-machine
```

### Development Checks

```powershell
# Run the high-signal local validation suite
npm test

# Start the local task pane server for manual checks
npm run dev-server
```

---

## Public URLs

| Surface | URL |
|---|---|
| Public site | `https://github.colorgeek.co/word-markdown-addin/` |
| Install page | `https://github.colorgeek.co/word-markdown-addin/install.html` |
| Manifest | `https://github.colorgeek.co/word-markdown-addin/manifest.store.xml` |
| Task pane | `https://github.colorgeek.co/word-markdown-addin/taskpane.html` |
| Support | `https://github.colorgeek.co/word-markdown-addin/support.html` |
| Privacy | `https://github.colorgeek.co/word-markdown-addin/privacy.html` |

---

## Documentation

| Topic | File |
|---|---|
| Online install | [docs/online-install.md](docs/online-install.md) |
| Online publishing | [docs/publish-online.md](docs/publish-online.md) |
| GitHub Pages deployment | [docs/github-pages.md](docs/github-pages.md) |
| Online smoke test | [docs/online-smoke-test.md](docs/online-smoke-test.md) |
| Windows single-machine path | [docs/single-machine.md](docs/single-machine.md) |
| Release checklist | [docs/release-checklist.md](docs/release-checklist.md) |
| Skill list | [docs/skill-list.md](docs/skill-list.md) |

---

## AI-Assisted Development

This project was developed with AI assistance.

| Model | Role |
|---|---|
| OpenAI Codex CLI | Implementation, documentation, UI/i18n review |

> **Disclaimer:** While the author has made every effort to review and validate
> the AI-generated code, no guarantee can be made regarding its correctness, security,
> or fitness for any particular purpose. Use at your own risk.

---

## License

[MIT License](LICENSE)
