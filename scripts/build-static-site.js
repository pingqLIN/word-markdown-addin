import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const args = process.argv.slice(2);

const readOption = (name) => {
  const directToken = `--${name}`;
  const prefix = `${directToken}=`;

  for (let index = 0; index < args.length; index += 1) {
    const token = args[index];

    if (token === directToken) {
      return args[index + 1];
    }

    if (token.startsWith(prefix)) {
      return token.slice(prefix.length);
    }
  }

  return undefined;
};

const rootDir = process.cwd();
const outputDir = path.resolve(rootDir, readOption("output") || "dist/site");
const manifestPath = path.resolve(
  rootDir,
  readOption("manifest") || "dist/manifest.store.xml",
);
const mappings = [
  ["src/taskpane.html", "taskpane.html"],
  ["src/js", "js"],
  ["src/styles", "styles"],
  ["src/lib", "lib"],
  ["src/locales", "locales"],
  ["assets", "assets"],
];

const buildSummary = [];
const publicSiteUrl = String(process.env.MANIFEST_HOST || "").trim();
const publicRepoUrl = String(
  process.env.SUPPORT_URL || "https://github.com/pingqLIN/word-markdown-addin",
).trim();
const publicManifestUrl = publicSiteUrl ? `${publicSiteUrl.replace(/\/+$/, "")}/manifest.store.xml` : "./manifest.store.xml";
const publicTaskpaneUrl = publicSiteUrl ? `${publicSiteUrl.replace(/\/+$/, "")}/taskpane.html` : "./taskpane.html";
const displayName = String(process.env.DISPLAY_NAME || "Word Markdown Companion").trim();
const marketplaceAssetId = String(process.env.MARKETPLACE_ASSET_ID || "").trim().toUpperCase();
const marketplaceAddinTitle = String(
  process.env.MARKETPLACE_ADDIN_TITLE || displayName,
).trim();
const marketplaceLanguage = String(process.env.MARKETPLACE_LINK_LANGUAGE || "en-US").trim();
const wordOnTheWebInstallLink =
  marketplaceAssetId && marketplaceAddinTitle
    ? `https://go.microsoft.com/fwlink/?linkid=2261098&templateid=${encodeURIComponent(
        marketplaceAssetId,
      )}&templatetitle=${encodeURIComponent(marketplaceAddinTitle)}`
    : "";
const installPageConfig = JSON.stringify({
  marketplaceAddinTitle,
  marketplaceAssetId,
  marketplaceLanguage,
  publicManifestUrl,
});

await rm(outputDir, { recursive: true, force: true });
await mkdir(outputDir, { recursive: true });

for (const [sourceRelativePath, targetRelativePath] of mappings) {
  const sourcePath = path.resolve(rootDir, sourceRelativePath);
  const targetPath = path.resolve(outputDir, targetRelativePath);

  await cp(sourcePath, targetPath, { recursive: true });
  buildSummary.push({
    source: sourceRelativePath,
    target: targetRelativePath,
  });
}

const landingPage = `<!DOCTYPE html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Word Markdown Companion</title>
    <style>
      :root {
        color-scheme: light;
        --bg: #f4f0e8;
        --panel: rgba(255, 255, 255, 0.86);
        --text: #1f1a16;
        --muted: #63584d;
        --accent: #b6552d;
        --accent-strong: #8d3d1c;
        --line: rgba(82, 63, 49, 0.15);
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        min-height: 100vh;
        font-family: "Segoe UI", "Noto Sans TC", sans-serif;
        color: var(--text);
        background:
          radial-gradient(circle at top left, rgba(230, 162, 98, 0.24), transparent 32rem),
          linear-gradient(180deg, #f8f4ec 0%, var(--bg) 100%);
      }
      main {
        max-width: 64rem;
        margin: 0 auto;
        padding: 3rem 1.5rem 4rem;
      }
      .panel {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 1.5rem;
        box-shadow: 0 18px 50px rgba(51, 35, 23, 0.08);
        padding: 2rem;
      }
      h1 {
        margin: 0 0 0.75rem;
        font-size: clamp(2rem, 5vw, 3.2rem);
        line-height: 1.05;
      }
      p {
        margin: 0 0 1rem;
        line-height: 1.7;
      }
      .muted { color: var(--muted); }
      .actions {
        display: flex;
        flex-wrap: wrap;
        gap: 0.75rem;
        margin: 1.5rem 0 2rem;
      }
      .button {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 2.8rem;
        padding: 0.75rem 1.1rem;
        border-radius: 999px;
        border: 1px solid transparent;
        font-weight: 600;
        text-decoration: none;
      }
      .button-primary {
        background: var(--accent);
        color: white;
      }
      .button-secondary {
        border-color: var(--line);
        color: var(--text);
        background: rgba(255, 255, 255, 0.72);
      }
      .grid {
        display: grid;
        gap: 1rem;
        grid-template-columns: repeat(auto-fit, minmax(15rem, 1fr));
        margin-top: 1.5rem;
      }
      .card {
        border: 1px solid var(--line);
        border-radius: 1rem;
        padding: 1rem;
        background: rgba(255, 255, 255, 0.6);
      }
      code {
        font-family: Consolas, "SFMono-Regular", monospace;
        font-size: 0.92em;
      }
      ul {
        margin: 0.4rem 0 0;
        padding-left: 1.2rem;
        line-height: 1.6;
      }
    </style>
  </head>
  <body>
    <main>
      <section class="panel">
        <p class="muted">GitHub Pages public host</p>
        <h1>Word Markdown Companion</h1>
        <p>這個站點提供 Word Office Add-in 的公開靜態資源。真正的 Office task pane 入口是 <code>taskpane.html</code>，一般瀏覽器訪客請從這裡下載 manifest 或查看 repo 文件，不要直接把本站當成完整 Web App。</p>
        <p class="muted">This public site hosts the Office add-in assets. The task pane is intended to run inside Microsoft Word, not as a standalone browser app.</p>
        <div class="actions">
          <a class="button button-primary" href="./install.html">快速安裝</a>
          <a class="button button-secondary" href="./manifest.store.xml">下載 Manifest</a>
          <a class="button button-secondary" href="./taskpane.html">直接開 taskpane.html</a>
          <a class="button button-secondary" href="${publicRepoUrl}">GitHub Repo</a>
        </div>
        <p class="muted">目前公開頁面與安裝入口都部署在同一個 GitHub Pages host，請優先從安裝頁開始。</p>
        <div class="grid">
          <article class="card">
            <strong>Pages 首頁</strong>
            <p class="muted"><code>${publicSiteUrl || "(build-time MANIFEST_HOST not set)"}</code></p>
          </article>
          <article class="card">
            <strong>Manifest</strong>
            <p class="muted"><code>${publicManifestUrl}</code></p>
          </article>
          <article class="card">
            <strong>Task pane</strong>
            <p class="muted"><code>${publicTaskpaneUrl}</code></p>
          </article>
        </div>
      </section>
    </main>
  </body>
</html>
`;
const installPage = `<!DOCTYPE html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${displayName} 安裝入口</title>
    <style>
      :root {
        color-scheme: light;
        --bg: #f4f0e8;
        --panel: rgba(255, 255, 255, 0.9);
        --panel-strong: #fffdfa;
        --text: #1f1a16;
        --muted: #63584d;
        --accent: #b6552d;
        --accent-strong: #8d3d1c;
        --line: rgba(82, 63, 49, 0.15);
        --soft: rgba(182, 85, 45, 0.1);
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        min-height: 100vh;
        font-family: "Segoe UI", "Noto Sans TC", sans-serif;
        color: var(--text);
        background:
          radial-gradient(circle at top left, rgba(230, 162, 98, 0.24), transparent 32rem),
          linear-gradient(180deg, #f8f4ec 0%, var(--bg) 100%);
      }
      main {
        max-width: 72rem;
        margin: 0 auto;
        padding: 3rem 1.5rem 4rem;
      }
      .panel {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 1.5rem;
        box-shadow: 0 18px 50px rgba(51, 35, 23, 0.08);
        padding: 2rem;
      }
      .stack {
        display: grid;
        gap: 1rem;
      }
      h1, h2, h3 {
        margin: 0;
      }
      h1 {
        font-size: clamp(2rem, 5vw, 3.2rem);
        line-height: 1.05;
      }
      h2 {
        font-size: 1.35rem;
      }
      p, li {
        line-height: 1.7;
      }
      .muted {
        color: var(--muted);
      }
      .eyebrow {
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0.16em;
        text-transform: uppercase;
        color: var(--accent-strong);
      }
      .actions {
        display: flex;
        flex-wrap: wrap;
        gap: 0.75rem;
      }
      .button {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 2.8rem;
        padding: 0.8rem 1.1rem;
        border-radius: 999px;
        border: 1px solid transparent;
        font-weight: 700;
        text-decoration: none;
        cursor: pointer;
        font: inherit;
      }
      .button-primary {
        background: var(--accent);
        color: white;
      }
      .button-secondary {
        border-color: var(--line);
        color: var(--text);
        background: rgba(255, 255, 255, 0.76);
      }
      .grid {
        display: grid;
        gap: 1rem;
        grid-template-columns: repeat(auto-fit, minmax(18rem, 1fr));
      }
      .card {
        border: 1px solid var(--line);
        border-radius: 1rem;
        padding: 1rem;
        background: rgba(255, 255, 255, 0.66);
      }
      .callout {
        border-left: 4px solid var(--accent);
        padding: 1rem 1rem 1rem 1.1rem;
        background: var(--soft);
        border-radius: 0.9rem;
      }
      .code-box {
        display: block;
        width: 100%;
        padding: 0.9rem 1rem;
        border-radius: 0.9rem;
        background: #1f1b19;
        color: #fff7ef;
        font-family: Consolas, "SFMono-Regular", monospace;
        font-size: 0.93rem;
        word-break: break-all;
      }
      ol {
        margin: 0;
        padding-left: 1.2rem;
      }
      .button-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.75rem;
        margin-top: 0.9rem;
      }
      @media (max-width: 640px) {
        main { padding: 2rem 1rem 3rem; }
        .panel { padding: 1.35rem; }
      }
    </style>
  </head>
  <body>
    <main class="stack">
      <section class="panel stack">
        <p class="eyebrow">Install</p>
        <h1>${displayName}</h1>
        <p>這頁整理目前這個公開版 add-in 的安裝入口。現階段可用的是 <strong>manifest 分發</strong> 與 <strong>tenant admin URL 部署</strong>；真正的 Microsoft 官方「click-and-run 一鍵安裝」只會在取得 Microsoft Marketplace asset ID 後出現。</p>
        <div class="grid">
          <article class="card">
            <h2>公開站點</h2>
            <p class="muted"><code>${publicSiteUrl || "(build-time MANIFEST_HOST not set)"}</code></p>
          </article>
          <article class="card">
            <h2>Manifest URL</h2>
            <p class="muted"><code id="manifest-url">${publicManifestUrl}</code></p>
          </article>
          <article class="card">
            <h2>Task pane URL</h2>
            <p class="muted"><code>${publicTaskpaneUrl}</code></p>
          </article>
        </div>
      </section>

      <section class="panel stack">
        <h2>目前最快可用的安裝方式</h2>
        <div class="grid">
          <article class="card stack">
            <h3>個人安裝</h3>
            <p>這是目前最直接的公開安裝方式。先下載 manifest，再在 Word 的 <code>My Add-ins</code> / <code>Upload My Add-in</code> 載入。</p>
            <div class="actions">
              <a class="button button-primary" href="./manifest.store.xml" download>下載 Manifest</a>
              <button id="copy-manifest-url" type="button" class="button button-secondary">複製 Manifest URL</button>
              <a class="button button-secondary" href="https://support.microsoft.com/en-us/office/view-manage-and-install-add-ins-for-excel-powerpoint-and-word-16278816-1948-4028-91e5-76dca5380f8d">官方安裝說明</a>
            </div>
          </article>

          <article class="card stack">
            <h3>Tenant Admin 安裝</h3>
            <p>如果你是 Microsoft 365 管理員，官方支援從 admin center 直接用 manifest URL 部署 add-in。</p>
            <div class="actions">
              <button id="copy-admin-manifest-url" type="button" class="button button-primary">複製部署用 URL</button>
              <a class="button button-secondary" href="https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide">Admin Center 官方文件</a>
            </div>
          </article>
        </div>
        <div class="callout">
          <strong>官方限制</strong>
          <p class="muted">目前這個 GitHub Pages 版本還沒有 Microsoft Marketplace asset ID，所以還不能生成 Microsoft 官方的 click-and-run 安裝連結。這不是本站限制，而是 Office Add-in 分發模型本身的限制。</p>
        </div>
      </section>

      <section class="panel stack">
        <h2>未來的真正一鍵安裝</h2>
        <p>一旦這個 add-in 取得 Microsoft Marketplace asset ID，這裡就會自動升級成官方的一鍵安裝頁，支援 Word on the web 與 Word desktop 的 click-and-run 連結。</p>
        <div id="marketplace-install" class="grid"></div>
      </section>

      <section class="panel stack">
        <h2>安裝後檢查</h2>
        <ol>
          <li>確認 Word ribbon 上出現 <code>Markdown</code> / <code>Markdown Tools</code>。</li>
          <li>確認 task pane 可開啟，且 URL 來源是 <code>${publicTaskpaneUrl}</code>。</li>
          <li>用 <code>Import .md</code> 匯入 Markdown，或用 <code>Export .md</code> 做回寫測試。</li>
        </ol>
        <div class="button-row">
          <a class="button button-secondary" href="./taskpane.html">檢視 taskpane.html</a>
          <a class="button button-secondary" href="${publicRepoUrl}">GitHub Repo</a>
        </div>
      </section>

      <script>
        const installConfig = ${installPageConfig};
        const manifestUrl = installConfig.publicManifestUrl;
        const copyButtons = [
          document.getElementById("copy-manifest-url"),
          document.getElementById("copy-admin-manifest-url"),
        ].filter(Boolean);

        const copyText = async (text, trigger) => {
          try {
            await navigator.clipboard.writeText(text);
            const previousLabel = trigger.textContent;
            trigger.textContent = "已複製";
            window.setTimeout(() => {
              trigger.textContent = previousLabel;
            }, 1400);
          } catch (error) {
            window.alert("複製失敗，請手動複製這個 URL:\\n" + text);
          }
        };

        for (const button of copyButtons) {
          button.addEventListener("click", () => copyText(manifestUrl, button));
        }

        const marketplaceContainer = document.getElementById("marketplace-install");
        if (installConfig.marketplaceAssetId) {
          const desktopUrl = () => {
            const correlationId =
              typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
                ? crypto.randomUUID()
                : "00000000-0000-4000-8000-000000000000";
            const encodedTitle = encodeURIComponent(installConfig.marketplaceAddinTitle);
            return "ms-word:https://api.addins.store.office.com/addinstemplate/"
              + encodeURIComponent(installConfig.marketplaceLanguage)
              + "/" + correlationId
              + "/" + encodeURIComponent(installConfig.marketplaceAssetId)
              + "/none/" + encodedTitle
              + ".docx?omexsrctype=1&isexternallink=1";
          };

          marketplaceContainer.innerHTML = [
            '<article class="card stack">',
            '  <h3>Word on the web</h3>',
            '  <p>官方 click-and-run 安裝連結。</p>',
            '  <div class="actions">',
            '    <a class="button button-primary" href="${wordOnTheWebInstallLink}">Open in Word on the web</a>',
            '  </div>',
            '</article>',
            '<article class="card stack">',
            '  <h3>Word Desktop</h3>',
            '  <p>官方 click-and-run 桌面版安裝連結。</p>',
            '  <div class="actions">',
            '    <button id="desktop-install" type="button" class="button button-primary">Open in Word Desktop</button>',
            '  </div>',
            '</article>',
          ].join("");

          const desktopButton = document.getElementById("desktop-install");
          if (desktopButton) {
            desktopButton.addEventListener("click", () => {
              window.location.href = desktopUrl();
            });
          }
        } else {
          marketplaceContainer.innerHTML = [
            '<article class="card stack">',
            '  <h3>尚未設定 Marketplace asset ID</h3>',
            '  <p>如果之後取得 Microsoft Marketplace asset ID，只要在建置時提供 <code>MARKETPLACE_ASSET_ID</code> 與 <code>MARKETPLACE_ADDIN_TITLE</code>，這頁就會自動生成官方一鍵安裝按鈕。</p>',
            '  <span class="code-box">MARKETPLACE_ASSET_ID=WA000000000\\nMARKETPLACE_ADDIN_TITLE=${marketplaceAddinTitle.replace(/"/g, "&quot;")}\\nMARKETPLACE_LINK_LANGUAGE=${marketplaceLanguage}</span>',
            '</article>',
          ].join("");
        }
      </script>
    </main>
  </body>
</html>
`;
await writeFile(path.resolve(outputDir, "index.html"), landingPage, "utf8");
await writeFile(path.resolve(outputDir, "install.html"), installPage, "utf8");
await writeFile(path.resolve(outputDir, ".nojekyll"), "", "utf8");
buildSummary.push({ source: "[generated]", target: "index.html" });
buildSummary.push({ source: "[generated]", target: "install.html" });
buildSummary.push({ source: "[generated]", target: ".nojekyll" });

try {
  await cp(manifestPath, path.resolve(outputDir, "manifest.store.xml"));
  buildSummary.push({
    source: path.relative(rootDir, manifestPath).replace(/\\/g, "/"),
    target: "manifest.store.xml",
  });
} catch (error) {
  if (error?.code !== "ENOENT") {
    throw error;
  }
}

const metadataPath = path.resolve(outputDir, "build-metadata.json");
await writeFile(
  metadataPath,
  JSON.stringify(
    {
      builtAt: new Date().toISOString(),
      outputDir: path.relative(rootDir, outputDir).replace(/\\/g, "/"),
      files: buildSummary,
    },
    null,
    2,
  ),
  "utf8",
);

const taskpanePath = path.resolve(outputDir, "taskpane.html");
const taskpaneContents = await readFile(taskpanePath, "utf8");
if (!taskpaneContents.includes('src="js/taskpane.js"')) {
  throw new Error(`Unexpected taskpane entrypoint in ${taskpanePath}`);
}

console.log(`Generated static site bundle at ${path.relative(rootDir, outputDir)}`);
