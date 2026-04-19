import http from "node:http";
import { stat, readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";

const PORT = 3000;
const ROOT = process.cwd();

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
  if (normalized.startsWith("src/")) {
    return path.join(ROOT, normalized);
  }
  return path.join(ROOT, "src", normalized);
};

const toEtag = (contents) =>
  `"${createHash("sha1").update(contents).digest("hex")}"`;

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "", `http://localhost:${PORT}`);
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
    console.error("Port 3000 is already in use. Stop the existing process or rerun with a different MANIFEST_HOST.");
    process.exit(1);
  }

  console.error("Server failed to start:", error);
  process.exit(1);
});

server.listen(PORT, () => {
  console.log(`Word Markdown Add-in dev server running: http://localhost:${PORT}`);
});
