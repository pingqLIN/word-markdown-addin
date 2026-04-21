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
          <a class="button button-primary" href="./manifest.store.xml">下載 Manifest</a>
          <a class="button button-secondary" href="./taskpane.html">直接開 taskpane.html</a>
          <a class="button button-secondary" href="${publicRepoUrl}">GitHub Repo</a>
        </div>
        <div class="grid">
          <article class="card">
            <strong>預設公開網址</strong>
            <p class="muted"><code>${publicSiteUrl || "(build-time MANIFEST_HOST not set)"}</code></p>
          </article>
          <article class="card">
            <strong>主要檔案</strong>
            <ul>
              <li><code>manifest.store.xml</code></li>
              <li><code>taskpane.html</code></li>
              <li><code>js/taskpane.js</code></li>
              <li><code>locales/*.json</code></li>
            </ul>
          </article>
          <article class="card">
            <strong>使用方式</strong>
            <ul>
              <li>先用 manifest 載入 Word add-in</li>
              <li>再在 Word 內開啟 task pane</li>
              <li>手動匯入或匯出 Markdown</li>
            </ul>
          </article>
        </div>
      </section>
    </main>
  </body>
</html>
`;
await writeFile(path.resolve(outputDir, "index.html"), landingPage, "utf8");
await writeFile(path.resolve(outputDir, ".nojekyll"), "", "utf8");
buildSummary.push({ source: "[generated]", target: "index.html" });
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
