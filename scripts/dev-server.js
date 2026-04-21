import http from "node:http";
import { stat, readFile, rm, appendFile, mkdir } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { DEFAULT_LOCAL_HOST, getPortFromHost, normalizeHost, readRuntimeHost } from "./runtime-config.js";

const ROOT = process.cwd();
const runtimeHost = normalizeHost(
  process.env.MANIFEST_HOST || (await readRuntimeHost(ROOT)) || DEFAULT_LOCAL_HOST,
);
const PORT = Number(process.env.PORT || getPortFromHost(runtimeHost));
const PENDING_MARKDOWN_PATH = path.join(
  ROOT,
  ".local",
  "pending-open.json",
);
const TASKPANE_LOG_PATH = path.join(ROOT, ".local", "taskpane.log");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".txt": "text/plain; charset=utf-8",
};

const resolvePath = (requestUrl) => {
  const safe = requestUrl === "/" ? "/taskpane.html" : requestUrl;
  const normalized = safe.replace(/^\/+/, "");
  if (normalized.startsWith("..")) {
    return null;
  }
  if (normalized === "manifest.xml") {
    return path.join(ROOT, "manifest.xml");
  }
  if (normalized.startsWith("assets/")) {
    return path.join(ROOT, normalized);
  }
  if (normalized.startsWith("src/")) {
    return path.join(ROOT, normalized);
  }
  return path.join(ROOT, "src", normalized);
};

const toEtag = (contents) =>
  `"${createHash("sha1").update(contents).digest("hex")}"`;

const writeJson = (res, statusCode, payload) => {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-cache",
    "Access-Control-Allow-Origin": "*",
  });
  res.end(JSON.stringify(payload));
};

const readPendingMarkdown = async () => {
  const contents = await readFile(PENDING_MARKDOWN_PATH, "utf8");
  return JSON.parse(contents);
};

const clearPendingMarkdown = async () => {
  await rm(PENDING_MARKDOWN_PATH, { force: true });
};

const appendTaskpaneLog = async (entry) => {
  const timestamp = new Date().toISOString();
  const line = `[${timestamp}] ${entry}\n`;
  await mkdir(path.dirname(TASKPANE_LOG_PATH), { recursive: true });
  await appendFile(TASKPANE_LOG_PATH, line, "utf8");
};

const readRequestBody = async (req) =>
  new Promise((resolve, reject) => {
    const chunks = [];

    req.on("data", (chunk) => {
      chunks.push(chunk);
    });
    req.on("end", () => {
      resolve(Buffer.concat(chunks).toString("utf8"));
    });
    req.on("error", reject);
  });

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "", runtimeHost);

    if (url.pathname === "/api/pending-markdown") {
      if (req.method === "DELETE") {
        await clearPendingMarkdown();
        res.writeHead(204, {
          "Cache-Control": "no-cache",
          "Access-Control-Allow-Origin": "*",
        });
        res.end();
        return;
      }

      if (req.method === "GET") {
        try {
          const pending = await readPendingMarkdown();
          writeJson(res, 200, pending);
          return;
        } catch (error) {
          if (error.code === "ENOENT") {
            res.writeHead(204, {
              "Cache-Control": "no-cache",
              "Access-Control-Allow-Origin": "*",
            });
            res.end();
            return;
          }

          throw error;
        }
      }

      res.writeHead(405, {
        Allow: "GET, DELETE",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("Method not allowed");
      return;
    }

    if (url.pathname === "/api/taskpane-log") {
      if (req.method !== "POST") {
        res.writeHead(405, {
          Allow: "POST",
          "Access-Control-Allow-Origin": "*",
        });
        res.end("Method not allowed");
        return;
      }

      const rawBody = await readRequestBody(req);
      const payload = rawBody ? JSON.parse(rawBody) : {};
      const message =
        payload && typeof payload.message === "string"
          ? payload.message.trim()
          : "";

      if (!message) {
        writeJson(res, 400, { error: "message is required" });
        return;
      }

      await appendTaskpaneLog(message);
      writeJson(res, 202, { ok: true });
      return;
    }

    const filePath = resolvePath(url.pathname);

    if (!filePath) {
      res.writeHead(400);
      res.end("Bad request");
      return;
    }

    const stats = await stat(filePath);
    if (!stats.isFile()) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contents = await readFile(filePath);
    const mimeType = MIME_TYPES[ext] || "application/octet-stream";
    const etag = toEtag(contents);
    const ifNoneMatch = req.headers["if-none-match"];

    if (ifNoneMatch === etag) {
      res.writeHead(304);
      res.end();
      return;
    }

    res.writeHead(200, {
      "Content-Type": mimeType,
      "Cache-Control": "no-cache",
      "Access-Control-Allow-Origin": "*",
      ETag: etag,
    });
    res.end(contents);
  } catch (error) {
    if (error.code === "ENOENT") {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    console.error("Server error:", error);
    res.writeHead(500);
    res.end("Internal error");
  }
});

server.on("error", (error) => {
  if (error.code === "EADDRINUSE") {
    console.error(`Port ${PORT} is already in use. Stop the existing process or rerun with a different MANIFEST_HOST.`);
    process.exit(1);
  }

  console.error("Server failed to start:", error);
  process.exit(1);
});

server.listen(PORT, () => {
  console.log(`Word Markdown Add-in dev server running: ${runtimeHost}`);
});
