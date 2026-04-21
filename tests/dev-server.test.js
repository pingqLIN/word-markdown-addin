import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import { getAvailablePort, repoRoot, waitFor } from "./shared.js";

const rootPath = new URL(".", repoRoot);
const pendingMarkdownPath = new URL("../.local/pending-open.json", import.meta.url);
const taskpaneLogPath = new URL("../.local/taskpane.log", import.meta.url);

let baseUrl;
let serverProcess;
let serverOutput = "";

before(async () => {
  const port = await getAvailablePort();
  baseUrl = `http://127.0.0.1:${port}`;

  await rm(pendingMarkdownPath, { force: true });
  await rm(taskpaneLogPath, { force: true });

  serverProcess = spawn(process.execPath, ["scripts/dev-server.js"], {
    cwd: rootPath,
    env: {
      ...process.env,
      MANIFEST_HOST: baseUrl,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  serverProcess.stdout.on("data", (chunk) => {
    serverOutput += chunk.toString();
  });

  serverProcess.stderr.on("data", (chunk) => {
    serverOutput += chunk.toString();
  });

  await waitFor(async () => {
    const response = await fetch(`${baseUrl}/taskpane.html`);
    return response.ok;
  });
});

after(async () => {
  await rm(pendingMarkdownPath, { force: true });
  await rm(taskpaneLogPath, { force: true });

  if (!serverProcess || serverProcess.killed) {
    return;
  }

  serverProcess.kill("SIGINT");
  await once(serverProcess, "close");
});

test("dev server serves taskpane content and static assets with cache validation", async () => {
  const response = await fetch(`${baseUrl}/taskpane.html`);
  assert.equal(response.status, 200);
  assert.match(await response.text(), /Word × Markdown/u);

  const etag = response.headers.get("etag");
  assert.ok(etag, `Expected ETag header. Output: ${serverOutput}`);

  const cachedResponse = await fetch(`${baseUrl}/taskpane.html`, {
    headers: {
      "If-None-Match": etag,
    },
  });
  assert.equal(cachedResponse.status, 304);

  const localeResponse = await fetch(`${baseUrl}/locales/zh-TW.json`);
  assert.equal(localeResponse.status, 200);
  assert.equal((await localeResponse.json()).meta.title, "Word × Markdown 助手");
});

test("dev server exposes pending-markdown and taskpane-log APIs", async () => {
  const missingPending = await fetch(`${baseUrl}/api/pending-markdown`);
  assert.equal(missingPending.status, 204);

  const pendingPayload = {
    fileName: "demo.md",
    markdown: "# Demo\n\nHello from the launcher.",
    createdAt: "2026-04-22T00:00:00.000Z",
  };
  await writeFile(
    pendingMarkdownPath,
    JSON.stringify(pendingPayload, null, 2),
    "utf8",
  );

  const pendingResponse = await fetch(`${baseUrl}/api/pending-markdown`);
  assert.equal(pendingResponse.status, 200);
  assert.deepEqual(await pendingResponse.json(), pendingPayload);

  const deleteResponse = await fetch(`${baseUrl}/api/pending-markdown`, {
    method: "DELETE",
  });
  assert.equal(deleteResponse.status, 204);

  const deletedPending = await fetch(`${baseUrl}/api/pending-markdown`);
  assert.equal(deletedPending.status, 204);

  const logResponse = await fetch(`${baseUrl}/api/taskpane-log`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: "test log entry",
    }),
  });
  assert.equal(logResponse.status, 202);

  const logContents = await readFile(taskpaneLogPath, "utf8");
  assert.match(logContents, /test log entry/);
});
