import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { runNodeCommand } from "./shared.js";

test("render-manifest generates an HTTPS store manifest without template placeholders", async () => {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "word-markdown-manifest-"));
  const outputPath = path.join(tempDir, "manifest.store.xml");

  try {
    await runNodeCommand(
      ["scripts/render-manifest.js", "--output", outputPath, "--require-https"],
      {
        env: {
          MANIFEST_HOST: "https://addin.example.test",
          SUPPORT_URL: "https://support.example.test/help",
          ADDIN_ID: "00000000-0000-0000-0000-000000000000",
          PROVIDER_NAME: "Example Provider",
          DISPLAY_NAME: "Example Add-in",
          ADDIN_DESCRIPTION: "Example description",
        },
      },
    );

    const contents = await readFile(outputPath, "utf8");

    assert.match(contents, /https:\/\/addin\.example\.test\/taskpane\.html/);
    assert.match(contents, /https:\/\/support\.example\.test\/help/);
    assert.match(contents, /<Id>00000000-0000-0000-0000-000000000000<\/Id>/);
    assert.equal(contents.includes("{{"), false, "Manifest should not contain unresolved placeholders.");
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
});

test("render-manifest rejects non-HTTPS store hosts", async () => {
  await assert.rejects(
    () =>
      runNodeCommand(
        ["scripts/render-manifest.js", "--output", "dist/test-manifest.xml", "--require-https"],
        {
          env: {
            MANIFEST_HOST: "http://localhost:3000",
            SUPPORT_URL: "https://support.example.test/help",
          },
        },
      ),
    /requires HTTPS/i,
  );
});
