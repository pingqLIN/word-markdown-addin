import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { runNodeCommand } from "./shared.js";

test("project-loop-state persists resumable batch metadata and appends events", async () => {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "word-markdown-loop-"));
  const statePath = path.join(tempDir, "project-loop-state.json");
  const eventsPath = path.join(tempDir, "project-loop-events.jsonl");

  try {
    await runNodeCommand([
      "scripts/project-loop-state.js",
      "start",
      "--state-file",
      statePath,
      "--events-file",
      eventsPath,
      "--mode",
      "pattern-b",
      "--deadline",
      "2026-04-22T14:52:00+08:00",
      "--active-batch",
      "Marketplace readiness",
      "--last-completed-checkpoint",
      "audit complete",
      "--next-intended-action",
      "add support and privacy pages",
    ]);

    await runNodeCommand([
      "scripts/project-loop-state.js",
      "checkpoint",
      "--state-file",
      statePath,
      "--events-file",
      eventsPath,
      "--deadline",
      "2026-04-22T14:52:00+08:00",
      "--active-batch",
      "Marketplace readiness",
      "--last-completed-checkpoint",
      "support and privacy pages generated",
      "--next-intended-action",
      "run tests and build",
    ]);

    const state = JSON.parse(await readFile(statePath, "utf8"));
    assert.equal(state.mode, "pattern-b");
    assert.equal(state.status, "active");
    assert.equal(state.activeBatch, "Marketplace readiness");
    assert.equal(state.deadline, "2026-04-22T14:52:00+08:00");
    assert.equal(state.lastCompletedCheckpoint, "support and privacy pages generated");
    assert.equal(state.nextIntendedAction, "run tests and build");
    assert.ok(state.startedAt);
    assert.ok(state.updatedAt);

    const eventLines = (await readFile(eventsPath, "utf8"))
      .trim()
      .split("\n")
      .map((line) => JSON.parse(line));
    assert.equal(eventLines.length, 2);
    assert.equal(eventLines[0].action, "start");
    assert.equal(eventLines[1].action, "checkpoint");
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
});
