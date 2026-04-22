import test from "node:test";

import { runNodeCommand } from "./shared.js";

const filesToCheck = [
  "scripts/render-manifest.js",
  "scripts/build-static-site.js",
  "scripts/dev-server.js",
  "scripts/project-loop-state.js",
  "scripts/setup-auto.js",
  "scripts/sideload.js",
  "src/js/taskpane.js",
];

for (const filePath of filesToCheck) {
  test(`node --check ${filePath}`, async () => {
    await runNodeCommand(["--check", filePath]);
  });
}
