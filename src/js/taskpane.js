const statusElement = document.getElementById("status");
const mdFileInput = document.getElementById("md-file");
const importButton = document.getElementById("import-btn");
const exportButton = document.getElementById("export-btn");
const exportPreview = document.getElementById("export-preview");
const downloadButton = document.getElementById("download-btn");
const dropZone = document.getElementById("drop-zone");
const mdAutoImportButton = document.getElementById("md-auto-import-btn");
const mdAutoImportHelper = document.getElementById("md-auto-import-helper");
const toolboxTitle = document.getElementById("toolbox-title");
const toolboxSummary = document.getElementById("toolbox-summary");
const panelLinks = Array.from(document.querySelectorAll("[data-panel-link]"));
const toolViews = Array.from(document.querySelectorAll("[data-tool-view]"));
const pendingMarkdownState = {
  fileName: "",
  markdown: "",
};
const toolViewMeta = {
  import: {
    title: "匯入 Markdown",
    summary: "把本機 Markdown 插入目前 Word 文件，或接手 launcher 剛交接的內容。",
  },
  export: {
    title: "匯出 Markdown",
    summary: "從目前文件抽出 Markdown，先預覽，再決定是否下載成檔案。",
  },
  format: {
    title: "格式整理",
    summary: "把純文字 Markdown 重新套用成 Word 版式，並處理待匯入的 launcher 內容。",
  },
};

const logTaskpaneEvent = async (message, extra = null) => {
  const detail =
    extra && typeof extra === "object"
      ? ` ${JSON.stringify(extra)}`
      : "";

  try {
    await fetch("/api/taskpane-log", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: `${message}${detail}`,
      }),
    });
  } catch (error) {
    console.warn("taskpane log failed", error);
  }
};

const requireMarkdownLibraries = () => {
  if (typeof window.marked?.parse !== "function") {
    throw new Error("marked 函式庫未載入，請檢查 src/lib/marked.min.js。");
  }

  if (typeof window.TurndownService !== "function") {
    throw new Error("turndown 函式庫未載入，請檢查 src/lib/turndown.min.js。");
  }
};

const turndownService = new TurndownService({
  codeBlockStyle: "fenced",
  headingStyle: "atx",
  bulletListMarker: "-",
  emDelimiter: "_",
  strongDelimiter: "**",
});

turndownService.remove("script");
turndownService.remove("style");

const createFilename = () => {
  const now = new Date();
  const safe = now
    .toISOString()
    .replace(/[:T]/g, "-")
    .replace(/\..+/, "")
    .replace(/Z$/, "");
  return `word-export-${safe}.md`;
};

const markdownToWord = (text) =>
  window.marked.parse(text || "", {
    breaks: true,
    gfm: true,
  });

const ensureOffice = () => {
  if (!window.Office) {
    throw new Error("Office.js 尚未載入完成");
  }
  if (!window.Office.context || !window.Office.context.document) {
    throw new Error("Office.context 不可用，請在 Word 任務窗格中執行。");
  }
};

const setStatus = (message) => {
  statusElement.textContent = message || "";
};

const activateToolView = (viewName) => {
  const targetViewName = toolViewMeta[viewName] ? viewName : "import";

  panelLinks.forEach((link) => {
    const isActive = link.dataset.panelLink === targetViewName;
    link.classList.toggle("is-active", isActive);
    link.setAttribute("aria-current", isActive ? "page" : "false");
  });

  toolViews.forEach((view) => {
    const isActive = view.dataset.toolView === targetViewName;
    view.hidden = !isActive;
    view.classList.toggle("is-active", isActive);
  });

  if (toolboxTitle) {
    toolboxTitle.textContent = toolViewMeta[targetViewName].title;
  }

  if (toolboxSummary) {
    toolboxSummary.textContent = toolViewMeta[targetViewName].summary;
  }
};

const setErrorStatus = (error) => {
  const message =
    error && typeof error.message === "string"
      ? error.message
      : "發生未預期錯誤，請稍後再試。";
  setStatus(message);
};

const requireOfficeRuntime = () => {
  if (
    typeof window.Office !== "object" ||
    typeof window.Office.onReady !== "function"
  ) {
    throw new Error(
      "Office.js 未載入完成，請確認任務窗格可連到 Microsoft hosted office.js。"
    );
  }
};

const toWordMarkdown = (markdown) => markdownToWord(markdown);

const isMarkdownLikeFilename = (url = "") =>
  /\.(md|markdown)$/i.test(String(url).split(/[?#]/)[0].trim());

const detectMarkdownDocument = () => {
  const url = window.Office.context?.document?.url || "";
  return isMarkdownLikeFilename(url);
};

const insertMarkdownIntoWord = async (markdown) => {
  requireMarkdownLibraries();
  ensureOffice();
  const html = toWordMarkdown(markdown);
  await logTaskpaneEvent("insertMarkdownIntoWord:start", {
    markdownLength: markdown.length,
  });

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertHtml(html, Word.InsertLocation.replace);
    await context.sync();
  });

  await logTaskpaneEvent("insertMarkdownIntoWord:success");
};

const formatExistingMarkdownDocument = async () => {
  requireMarkdownLibraries();
  ensureOffice();

  const markdown = await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text || "";
  });

  if (!markdown.trim()) {
    return "文件內目前沒有可轉換的內容。";
  }

  const html = toWordMarkdown(markdown);
  await Word.run(async (context) => {
    const body = context.document.body;
    body.clear();
    body.insertHtml(html, Word.InsertLocation.end);
    await context.sync();
  });
  return "已將目前文件按 Markdown 格式寫回 Word。";
};

const exportWordToMarkdown = async () => {
  requireMarkdownLibraries();
  ensureOffice();

  return Word.run(async (context) => {
    const body = context.document.body;
    let rawHtml = "";

    try {
      const htmlResult = body.getHtml();
      await context.sync();
      rawHtml = htmlResult.value || "";
    } catch (error) {
      body.load("text");
      await context.sync();
      rawHtml = body.text ? `<p>${String(body.text)}</p>` : "";
    }

    if (!rawHtml.trim()) {
      return "";
    }

    return turndownService.turndown(rawHtml);
  });
};

const setDownloadButtonEnabled = (enabled) => {
  downloadButton.disabled = !enabled;
};

const setAutoImportState = (isVisible, helperText = "", buttonText = "") => {
  if (mdAutoImportButton) {
    mdAutoImportButton.hidden = !isVisible;
    if (buttonText) {
      mdAutoImportButton.textContent = buttonText;
    }
  }

  if (mdAutoImportHelper) {
    mdAutoImportHelper.textContent = helperText || "";
    mdAutoImportHelper.classList.toggle("visible", Boolean(helperText));
  }
};

const getDocumentText = async () =>
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text || "";
  });

const fetchPendingMarkdown = async () => {
  const response = await fetch("/api/pending-markdown", {
    method: "GET",
    cache: "no-store",
  });

  if (response.status === 204) {
    return null;
  }

  if (!response.ok) {
    throw new Error("無法讀取待匯入的 Markdown 檔案。");
  }

  const payload = await response.json();
  if (!payload || typeof payload.markdown !== "string") {
    return null;
  }

  return payload;
};

const enableDocumentAutoShowTaskpane = async () => {
  ensureOffice();

  const settings = window.Office.context?.document?.settings;
  if (!settings || typeof settings.set !== "function") {
    return;
  }

  settings.set("Office.AutoShowTaskpaneWithDocument", true);

  await new Promise((resolvePromise, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolvePromise();
        return;
      }

      reject(result.error || new Error("無法儲存自動開啟 taskpane 設定。"));
    });
  });
};

const clearPendingMarkdown = async () => {
  await fetch("/api/pending-markdown", {
    method: "DELETE",
    cache: "no-store",
  });
};

const triggerDownload = (content) => {
  const blob = new Blob(["\uFEFF", content || ""], {
    type: "text/markdown;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = createFilename();
  link.rel = "noopener";
  link.click();
  URL.revokeObjectURL(url);
};

const exportMarkdownFile = async () => {
  try {
    setStatus("正在匯出…");
    const markdown = await exportWordToMarkdown();
    exportPreview.value = markdown;
    setDownloadButtonEnabled(Boolean(markdown));
    setStatus(markdown ? "匯出成功，可點「下載為 Markdown 檔」儲存。" : "文件目前為空，沒有可匯出的內容。");
    return markdown;
  } catch (error) {
    setErrorStatus(error);
    throw error;
  }
};

const handleDownload = async () => {
  try {
    if (!exportPreview.value) {
      setStatus("請先匯出 Markdown 後再下載。");
      return;
    }
    triggerDownload(exportPreview.value);
  } catch (error) {
    setErrorStatus(error);
  }
};

const handleAutoImport = async () => {
  try {
    setStatus("正在依據現有文件內容建立 Markdown 格式…");
    const message = await formatExistingMarkdownDocument();
    setStatus(message);
  } catch (error) {
    setErrorStatus(error);
  }
};

const handlePendingMarkdownImport = async () => {
  try {
    if (!pendingMarkdownState.markdown) {
      setStatus("目前沒有待匯入的 Markdown 檔案。");
      await logTaskpaneEvent("handlePendingMarkdownImport:no-pending-markdown");
      return;
    }

    setStatus(`正在匯入 ${pendingMarkdownState.fileName || "Markdown 檔案"}…`);
    await logTaskpaneEvent("handlePendingMarkdownImport:start", {
      fileName: pendingMarkdownState.fileName || "",
      markdownLength: pendingMarkdownState.markdown.length,
    });
    await insertMarkdownIntoWord(pendingMarkdownState.markdown);
    await enableDocumentAutoShowTaskpane();
    await clearPendingMarkdown();
    await logTaskpaneEvent("handlePendingMarkdownImport:cleared-pending");

    pendingMarkdownState.fileName = "";
    pendingMarkdownState.markdown = "";
    setAutoImportState(false);
    setStatus("已匯入剛開啟的 Markdown 檔案。");
    await logTaskpaneEvent("handlePendingMarkdownImport:success");
  } catch (error) {
    setErrorStatus(error);
    await logTaskpaneEvent("handlePendingMarkdownImport:error", {
      message: error?.message || String(error),
    });
  }
};

const handleMdFile = (file) => {
  if (!file) {
    setStatus("請先選擇 .md 檔案。");
    return;
  }

  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext !== "md" && ext !== "markdown") {
    setStatus("僅支援 .md / .markdown 檔案。");
    return;
  }

  const reader = new FileReader();
  reader.onload = async (event) => {
    try {
      const markdown = event.target?.result?.toString() || "";
      await insertMarkdownIntoWord(markdown);
      await enableDocumentAutoShowTaskpane();
      setStatus("已匯入 Markdown。");
    } catch (error) {
      setErrorStatus(error);
    }
  };
  reader.onerror = () => {
    setStatus("無法讀取檔案。");
  };
  reader.readAsText(file);
};

const onDropFiles = (event) => {
  event.preventDefault();
  dropZone.classList.remove("active");
  const files = event.dataTransfer && event.dataTransfer.files;
  if (!files || !files.length) {
    setStatus("未偵測到可用檔案。");
    return;
  }
  handleMdFile(files[0]);
};

requireOfficeRuntime();
void logTaskpaneEvent("taskpane-script-loaded");

Office.onReady(() => {
  try {
    requireMarkdownLibraries();
  } catch (error) {
    setErrorStatus(error);
    setDownloadButtonEnabled(false);
    void logTaskpaneEvent("Office.onReady:library-check-failed", {
      message: error?.message || String(error),
    });
    return;
  }

  void logTaskpaneEvent("Office.onReady:ready");

  const mdDetected = detectMarkdownDocument();
  activateToolView(mdDetected ? "format" : "import");

  panelLinks.forEach((link) => {
    link.addEventListener("click", (event) => {
      event.preventDefault();
      const targetView = link.dataset.panelLink || "import";
      activateToolView(targetView);
      document.getElementById("workspace-toolbox")?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    });
  });

  importButton.addEventListener("click", () => {
    mdFileInput.click();
  });
  if (mdAutoImportButton) {
    mdAutoImportButton.addEventListener("click", () => {
      if (pendingMarkdownState.markdown) {
        handlePendingMarkdownImport();
        return;
      }

      handleAutoImport();
    });
  }

  mdFileInput.addEventListener("change", () => {
    const file = mdFileInput.files && mdFileInput.files[0];
    if (file) {
      handleMdFile(file);
    }
  });
  exportButton.addEventListener("click", exportMarkdownFile);
  downloadButton.addEventListener("click", handleDownload);
  dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZone.classList.add("active");
  });
  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("active");
  });
  dropZone.addEventListener("drop", onDropFiles);
  setStatus("初始化完成，可匯入或匯出 Markdown。");
  setDownloadButtonEnabled(false);

  void (async () => {
    try {
      await logTaskpaneEvent("pending-check:start");
      const pending = await fetchPendingMarkdown();
      if (pending) {
        pendingMarkdownState.fileName = pending.fileName || "";
        pendingMarkdownState.markdown = pending.markdown || "";
        await logTaskpaneEvent("pending-check:found", {
          fileName: pendingMarkdownState.fileName,
          markdownLength: pendingMarkdownState.markdown.length,
        });

        const existingText = await getDocumentText();
        await logTaskpaneEvent("pending-check:document-text-loaded", {
          textLength: existingText.length,
        });
        setAutoImportState(
          true,
          pendingMarkdownState.fileName
            ? `偵測到剛由 launcher 交接的 Markdown：${pendingMarkdownState.fileName}`
            : "偵測到剛由 launcher 交接的 Markdown 檔案。",
          "匯入剛開啟的 Markdown 檔",
        );

        if (!existingText.trim()) {
          await logTaskpaneEvent("pending-check:auto-import-eligible");
          await handlePendingMarkdownImport();
          return;
        }

        setStatus("已偵測到待匯入 Markdown 檔，可點按鈕插入到目前文件。");
        await logTaskpaneEvent("pending-check:manual-import-required");
        return;
      }

      await logTaskpaneEvent("pending-check:none");
      const mdHelperText = mdDetected
        ? "偵測到 Markdown 文件，點擊可將目前純文字內容直接轉為 Word 格式。"
        : "";
      setAutoImportState(mdDetected, mdHelperText, "將目前文件格式化為 Markdown");
    } catch (error) {
      const mdHelperText = mdDetected
        ? "偵測到 Markdown 文件，點擊可將目前純文字內容直接轉為 Word 格式。"
        : "";
      setAutoImportState(mdDetected, mdHelperText, "將目前文件格式化為 Markdown");
      console.error(error);
      await logTaskpaneEvent("pending-check:error", {
        message: error?.message || String(error),
      });
    }
  })();
});
