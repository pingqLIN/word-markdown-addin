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
const localeButtons = Array.from(document.querySelectorAll("[data-locale-switch]"));
const themeButtons = Array.from(document.querySelectorAll("[data-theme-switch]"));

const supportedLocales = ["zh-TW", "en-US"];
const defaultLocale = "zh-TW";
const supportedThemes = ["warm", "dark"];
const defaultTheme = "warm";
const themeStorageKey = "wordMarkdownTheme";
const pendingMarkdownState = {
  fileName: "",
  markdown: "",
};

let activeToolView = "import";
let currentLocale = defaultLocale;
let currentTheme = defaultTheme;
let localeMessages = {};

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

const getTranslation = (path, fallback = "") => {
  const resolved = path
    .split(".")
    .reduce(
      (value, key) => (value && typeof value === "object" ? value[key] : undefined),
      localeMessages,
    );

  return typeof resolved === "string" ? resolved : fallback;
};

const interpolate = (template, values = {}) =>
  String(template).replace(/\{(\w+)\}/g, (_, key) => values[key] ?? "");

const t = (path, values = {}, fallback = "") =>
  interpolate(getTranslation(path, fallback), values);

const detectPreferredLocale = () => {
  const candidates = [
    String(window.Office?.context?.displayLanguage || "").trim(),
    String(window.navigator?.language || "").trim(),
  ];

  for (const locale of candidates) {
    if (!locale) {
      continue;
    }

    const directMatch = supportedLocales.find(
      (supportedLocale) => supportedLocale.toLowerCase() === locale.toLowerCase(),
    );
    if (directMatch) {
      return directMatch;
    }

    if (locale.toLowerCase().startsWith("zh")) {
      return "zh-TW";
    }

    if (locale.toLowerCase().startsWith("en")) {
      return "en-US";
    }
  }

  return defaultLocale;
};

const detectPreferredTheme = () => {
  try {
    const savedTheme = String(window.localStorage?.getItem(themeStorageKey) || "").trim();
    if (savedTheme === "microsoft") {
      return "dark";
    }
    if (supportedThemes.includes(savedTheme)) {
      return savedTheme;
    }
  } catch {
    return defaultTheme;
  }

  return defaultTheme;
};

const applyTheme = (theme, { persist = true } = {}) => {
  const normalizedTheme = supportedThemes.includes(theme) ? theme : defaultTheme;
  currentTheme = normalizedTheme;
  document.documentElement.dataset.theme = normalizedTheme;

  if (persist) {
    try {
      window.localStorage?.setItem(themeStorageKey, normalizedTheme);
    } catch {
      // Ignore storage failures. The active theme still applies for this session.
    }
  }

  themeButtons.forEach((button) => {
    const isActive = button.dataset.themeSwitch === normalizedTheme;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });
};

const loadLocaleMessages = async (locale) => {
  const normalizedLocale = supportedLocales.includes(locale) ? locale : defaultLocale;
  const response = await fetch(`locales/${normalizedLocale}.json`, {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(`Failed to load locale: ${normalizedLocale}`);
  }

  localeMessages = await response.json();
  currentLocale = normalizedLocale;
  document.documentElement.lang = normalizedLocale === "zh-TW" ? "zh-Hant" : normalizedLocale;
};

const applyTranslations = () => {
  document.querySelectorAll("[data-i18n]").forEach((element) => {
    const key = element.dataset.i18n;
    if (!key) {
      return;
    }

    element.textContent = getTranslation(key, element.textContent);
  });

  document.querySelectorAll("[data-i18n-html]").forEach((element) => {
    const key = element.dataset.i18nHtml;
    if (!key) {
      return;
    }

    element.innerHTML = getTranslation(key, element.innerHTML);
  });

  document.querySelectorAll("[data-i18n-placeholder]").forEach((element) => {
    const key = element.dataset.i18nPlaceholder;
    if (!key) {
      return;
    }

    element.setAttribute(
      "placeholder",
      getTranslation(key, element.getAttribute("placeholder") || ""),
    );
  });

  localeButtons.forEach((button) => {
    const isActive = button.dataset.localeSwitch === currentLocale;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });

  themeButtons.forEach((button) => {
    const themeName = button.dataset.themeSwitch;
    if (!themeName) {
      return;
    }

    const themeLabel = getTranslation(`theme.${themeName}`, themeName);
    button.setAttribute("aria-label", themeLabel);
    button.setAttribute("title", themeLabel);
  });
};

const getToolViewMeta = (viewName) => ({
  title: t(`toolbox.${viewName}.title`),
  summary: t(`toolbox.${viewName}.summary`),
});

const requireMarkdownLibraries = () => {
  if (typeof window.marked?.parse !== "function") {
    throw new Error("marked library is unavailable.");
  }

  if (typeof window.TurndownService !== "function") {
    throw new Error("turndown library is unavailable.");
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
    throw new Error(t("status.officeNotReady"));
  }
  if (!window.Office.context || !window.Office.context.document) {
    throw new Error(t("status.officeUnavailable"));
  }
};

const setStatus = (message) => {
  statusElement.textContent = message || "";
};

const setErrorStatus = (error) => {
  const message =
    error && typeof error.message === "string"
      ? error.message
      : t("status.unexpectedError");
  setStatus(message);
};

const activateToolView = (viewName) => {
  const targetViewName = ["import", "export", "format"].includes(viewName)
    ? viewName
    : "import";
  const toolViewMeta = getToolViewMeta(targetViewName);
  activeToolView = targetViewName;

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
    toolboxTitle.textContent = toolViewMeta.title;
  }

  if (toolboxSummary) {
    toolboxSummary.textContent = toolViewMeta.summary;
  }
};

const requireOfficeRuntime = () => {
  if (
    typeof window.Office !== "object" ||
    typeof window.Office.onReady !== "function"
  ) {
    throw new Error(t("status.officeNotReady"));
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
    return t("status.noConvertibleContent");
  }

  const html = toWordMarkdown(markdown);
  await Word.run(async (context) => {
    const body = context.document.body;
    body.clear();
    body.insertHtml(html, Word.InsertLocation.end);
    await context.sync();
  });

  return t("status.formatSuccess");
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
    throw new Error(t("status.cannotReadPending"));
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

      reject(result.error || new Error(t("status.saveAutoShowFailed")));
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
    setStatus(t("status.exporting"));
    const markdown = await exportWordToMarkdown();
    exportPreview.value = markdown;
    setDownloadButtonEnabled(Boolean(markdown));
    setStatus(markdown ? t("status.downloadReady") : t("status.documentEmpty"));
    return markdown;
  } catch (error) {
    setErrorStatus(error);
    throw error;
  }
};

const handleDownload = async () => {
  try {
    if (!exportPreview.value) {
      setStatus(t("status.downloadAfterExport"));
      return;
    }
    triggerDownload(exportPreview.value);
  } catch (error) {
    setErrorStatus(error);
  }
};

const handleAutoImport = async () => {
  try {
    setStatus(t("status.formattingExistingDocument"));
    const message = await formatExistingMarkdownDocument();
    setStatus(message);
  } catch (error) {
    setErrorStatus(error);
  }
};

const handlePendingMarkdownImport = async () => {
  try {
    if (!pendingMarkdownState.markdown) {
      setStatus(t("status.pendingFileMissing"));
      await logTaskpaneEvent("handlePendingMarkdownImport:no-pending-markdown");
      return;
    }

    setStatus(t("status.importing"));
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
    setStatus(t("status.pendingImportSuccess"));
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
    setStatus(t("status.selectMarkdownFirst"));
    return;
  }

  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext !== "md" && ext !== "markdown") {
    setStatus(t("status.onlyMarkdownSupported"));
    return;
  }

  const reader = new FileReader();
  reader.onload = async (event) => {
    try {
      const markdown = event.target?.result?.toString() || "";
      await insertMarkdownIntoWord(markdown);
      await enableDocumentAutoShowTaskpane();
      setStatus(t("status.importSuccess"));
    } catch (error) {
      setErrorStatus(error);
    }
  };
  reader.onerror = () => {
    setStatus(t("status.fileReadFailed"));
  };
  reader.readAsText(file);
};

const onDropFiles = (event) => {
  event.preventDefault();
  dropZone.classList.remove("active");
  const files = event.dataTransfer && event.dataTransfer.files;
  if (!files || !files.length) {
    setStatus(t("status.dropzoneNoFile"));
    return;
  }
  handleMdFile(files[0]);
};

applyTheme(detectPreferredTheme(), {
  persist: false,
});

requireOfficeRuntime();
void logTaskpaneEvent("taskpane-script-loaded");

Office.onReady(async () => {
  try {
    await loadLocaleMessages(detectPreferredLocale());
  } catch {
    await loadLocaleMessages(defaultLocale);
  }

  applyTranslations();

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

  localeButtons.forEach((button) => {
    button.addEventListener("click", async () => {
      const nextLocale = button.dataset.localeSwitch;
      if (!nextLocale || nextLocale === currentLocale) {
        return;
      }

      try {
        await loadLocaleMessages(nextLocale);
        applyTranslations();
        activateToolView(activeToolView);
        const mdHelperText = mdDetected
          ? t("status.markdownDocumentDetected")
          : "";
        setAutoImportState(mdDetected, mdHelperText, t("format.primaryAction"));
        setStatus(t("status.ready"));
      } catch (error) {
        setErrorStatus(error);
      }
    });
  });

  themeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const nextTheme = button.dataset.themeSwitch;
      if (!nextTheme || nextTheme === currentTheme) {
        return;
      }

      applyTheme(nextTheme);
    });
  });

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
  setStatus(t("status.ready"));
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
            ? t("status.pendingImportedFile", { fileName: pendingMarkdownState.fileName })
            : t("status.pendingImportedGeneric"),
          t("format.primaryAction"),
        );

        if (!existingText.trim()) {
          await logTaskpaneEvent("pending-check:auto-import-eligible");
          await handlePendingMarkdownImport();
          return;
        }

        setStatus(t("status.pendingImportDetected"));
        await logTaskpaneEvent("pending-check:manual-import-required");
        return;
      }

      await logTaskpaneEvent("pending-check:none");
      const mdHelperText = mdDetected
        ? t("status.markdownDocumentDetected")
        : "";
      setAutoImportState(mdDetected, mdHelperText, t("format.primaryAction"));
    } catch (error) {
      const mdHelperText = mdDetected
        ? t("status.markdownDocumentDetected")
        : "";
      setAutoImportState(mdDetected, mdHelperText, t("format.primaryAction"));
      console.error(error);
      await logTaskpaneEvent("pending-check:error", {
        message: error?.message || String(error),
      });
    }
  })();
});
