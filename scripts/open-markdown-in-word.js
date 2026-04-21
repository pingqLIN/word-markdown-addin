import { appendFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import { resolve, basename, extname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { DEFAULT_LOCAL_HOST, normalizeHost, readRuntimeHost } from "./runtime-config.js";

const defaultWordPath = join(
  process.env.ProgramFiles || "C:\\Program Files",
  "Microsoft Office",
  "root",
  "Office16",
  "WINWORD.EXE",
);

const markdownPathArg = process.argv[2];
const wordPath = process.env.WORD_PATH || defaultWordPath;

if (!markdownPathArg) {
  console.error("Markdown path is required.");
  process.exit(1);
}

const repoRoot = resolve(fileURLToPath(new URL("..", import.meta.url)));
const runtimeHost = normalizeHost(
  (await readRuntimeHost(repoRoot)) || process.env.MANIFEST_HOST || DEFAULT_LOCAL_HOST,
);
const pendingDirectory = join(repoRoot, ".local");
const pendingPath = join(pendingDirectory, "pending-open.json");
const launcherLogPath = join(pendingDirectory, "launcher.log");
const devServerUrl = `${runtimeHost}/taskpane.html`;
const compatibilityUrl = `${runtimeHost}/api/pending-markdown`;

const wait = (ms) => new Promise((resolveWait) => setTimeout(resolveWait, ms));

const logEvent = async (message) => {
  await mkdir(pendingDirectory, { recursive: true });
  await appendFile(
    launcherLogPath,
    `[${new Date().toISOString()}] ${message}\n`,
    "utf8",
  );
};

const isEndpointReady = async () => {
  try {
    const pageResponse = await fetch(devServerUrl, {
      method: "GET",
      cache: "no-store",
    });

    if (!pageResponse.ok) {
      return false;
    }

    const compatibilityResponse = await fetch(compatibilityUrl, {
      method: "GET",
      cache: "no-store",
    });

    return compatibilityResponse.status === 200 || compatibilityResponse.status === 204;
  } catch {
    return false;
  }
};

const startDetached = (command, args, cwd) => {
  const child = spawn(command, args, {
    cwd,
    detached: true,
    stdio: "ignore",
  });
  child.unref();
};

try {
  const inputPath = resolve(markdownPathArg);
  const extension = extname(inputPath).toLowerCase();

  await logEvent(`launcher invoked for ${inputPath}`);

  if (extension !== ".md" && extension !== ".markdown") {
    startDetached(wordPath, [inputPath], repoRoot);
    await logEvent(`non-markdown file passed through to Word: ${inputPath}`);
    process.exit(0);
  }

  const markdown = await readFile(inputPath, "utf8");

  await mkdir(pendingDirectory, { recursive: true });
  await writeFile(
    pendingPath,
    JSON.stringify({
      fileName: basename(inputPath),
      fullPath: inputPath,
      markdown,
      createdAt: new Date().toISOString(),
    }),
    "utf8",
  );
  await logEvent(`pending markdown written: ${pendingPath}`);

  if (!await isEndpointReady()) {
    await logEvent("dev server not ready; starting detached server");
    startDetached(process.execPath, ["scripts/dev-server.js"], repoRoot);
    await wait(2500);
  } else {
    await logEvent("dev server already ready");
  }

  startDetached(wordPath, ["/w"], repoRoot);
  await logEvent(`word launched: ${wordPath}`);
} catch (error) {
  try {
    await logEvent(`launcher error: ${error?.message || error}`);
  } catch {
    // Ignore logging failures while reporting the root error.
  }
  console.error(error?.message || error);
  process.exit(1);
}
