const statusElement = document.getElementById("status");
const mdFileInput = document.getElementById("md-file");
const importButton = document.getElementById("import-btn");
const exportButton = document.getElementById("export-btn");
const exportPreview = document.getElementById("export-preview");
const downloadButton = document.getElementById("download-btn");
const dropZone = document.getElementById("drop-zone");
const mdAutoImportButton = document.getElementById("md-auto-import-btn");
const mdAutoImportHelper = document.getElementById("md-auto-import-helper");

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
      "Office.js 未載入完成，請確認任務窗格頁面的 `lib/office.js` 可正確讀取。"
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

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertHtml(html, Word.InsertLocation.replace);
    await context.sync();
  });
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

const setAutoImportState = (isVisible, helperText = "") => {
  if (mdAutoImportButton) {
    mdAutoImportButton.hidden = !isVisible;
  }

  if (mdAutoImportHelper) {
    mdAutoImportHelper.textContent = helperText || "";
    mdAutoImportHelper.classList.toggle("visible", Boolean(helperText));
  }
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

Office.onReady(() => {
  try {
    requireMarkdownLibraries();
  } catch (error) {
    setErrorStatus(error);
    setDownloadButtonEnabled(false);
    return;
  }

  const mdDetected = detectMarkdownDocument();
  const mdHelperText = mdDetected
    ? "偵測到 Markdown 文件，點擊可將目前純文字內容直接轉為 Word 格式。"
    : "";
  setAutoImportState(mdDetected, mdHelperText);

  importButton.addEventListener("click", () => {
    mdFileInput.click();
  });
  if (mdAutoImportButton) {
    mdAutoImportButton.addEventListener("click", handleAutoImport);
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
});
