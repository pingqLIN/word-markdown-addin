import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { runNodeCommand } from "./shared.js";

test("build-static-site generates a hostable site bundle with expected files", async () => {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "word-markdown-site-"));
  const outputDir = path.join(tempDir, "site");

  try {
    const requiredFiles = [
      "index.html",
      "install.html",
      ".nojekyll",
      "manifest.store.xml",
      "taskpane.html",
      "js/taskpane.js",
      "styles/taskpane.css",
      "lib/marked.min.js",
      "lib/turndown.min.js",
      "locales/zh-TW.json",
      "locales/en-US.json",
      "assets/icon-16.png",
      "assets/icon-32.png",
      "assets/icon-80.png",
      "build-metadata.json",
    ];

    const manifestPath = path.join(tempDir, "manifest.store.xml");
    await rm(manifestPath, { force: true });
    await writeFile(manifestPath, "<manifest>demo</manifest>", "utf8");

    await runNodeCommand([
      "scripts/build-static-site.js",
      "--output",
      outputDir,
      "--manifest",
      manifestPath,
    ], {
      env: {
        DISPLAY_NAME: "Word Markdown Companion",
        MANIFEST_HOST: "https://github.colorgeek.co/word-markdown-addin",
        MARKETPLACE_ADDIN_TITLE: "Word Markdown Companion",
        MARKETPLACE_ASSET_ID: "WA200006278",
        MARKETPLACE_LINK_LANGUAGE: "en-US",
        SUPPORT_URL: "https://github.com/pingqLIN/word-markdown-addin",
      },
    });

    for (const relativePath of requiredFiles) {
      const candidatePath = path.join(outputDir, relativePath);
      const candidateStats = await stat(candidatePath);
      assert.equal(candidateStats.isFile(), true, `${relativePath} should exist in the site bundle.`);
    }

    const landingContents = await readFile(path.join(outputDir, "index.html"), "utf8");
    assert.match(landingContents, /Word Markdown Companion/);
    assert.match(landingContents, /install\.html/);
    assert.match(landingContents, /manifest\.store\.xml/);

    const installContents = await readFile(path.join(outputDir, "install.html"), "utf8");
    assert.match(
      installContents,
      /https:\/\/github\.colorgeek\.co\/word-markdown-addin\/manifest\.store\.xml/,
    );
    assert.match(installContents, /Open in Word on the web/);
    assert.match(installContents, /WA200006278/);

    const taskpaneContents = await readFile(path.join(outputDir, "taskpane.html"), "utf8");
    assert.match(taskpaneContents, /js\/taskpane\.js/);

    const metadata = JSON.parse(
      await readFile(path.join(outputDir, "build-metadata.json"), "utf8"),
    );
    assert.equal(metadata.outputDir, outputDir.replace(/\\/g, "/"));
    assert.ok(Array.isArray(metadata.files));
    assert.ok(metadata.files.length >= 9);
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
});
