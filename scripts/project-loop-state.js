import { appendFile, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const args = process.argv.slice(2);
const supportedActions = new Set(["start", "checkpoint", "complete"]);

const action = supportedActions.has(args[0]) ? args[0] : "start";

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
const localDir = path.resolve(rootDir, ".local");
const statePath = path.resolve(
  rootDir,
  readOption("state-file") || ".local/project-development-loop-active.json",
);
const eventsPath = path.resolve(
  rootDir,
  readOption("events-file") || ".local/project-development-loop-events.jsonl",
);

const readState = async () => {
  try {
    const contents = await readFile(statePath, "utf8");
    const parsed = JSON.parse(contents);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    if (error?.code === "ENOENT") {
      return {};
    }

    throw error;
  }
};

const now = new Date().toISOString();
const previousState = await readState();
const status = action === "complete" ? "completed" : "active";
const nextState = {
  schemaVersion: 1,
  projectRoot: rootDir,
  statePath: path.relative(rootDir, statePath).replace(/\\/g, "/"),
  eventsPath: path.relative(rootDir, eventsPath).replace(/\\/g, "/"),
  status,
  mode: readOption("mode") || previousState.mode || "pattern-b",
  activeBatch: readOption("active-batch") || previousState.activeBatch || "",
  deadline: readOption("deadline") || previousState.deadline || "",
  startedAt: previousState.startedAt || now,
  updatedAt: now,
  lastCompletedCheckpoint:
    readOption("last-completed-checkpoint") ||
    previousState.lastCompletedCheckpoint ||
    "",
  nextIntendedAction:
    readOption("next-intended-action") ||
    previousState.nextIntendedAction ||
    "",
  note: readOption("note") || "",
};

if (!nextState.activeBatch) {
  throw new Error("Missing required --active-batch value.");
}

if (!nextState.deadline) {
  throw new Error("Missing required --deadline value.");
}

await mkdir(localDir, { recursive: true });
await writeFile(statePath, JSON.stringify(nextState, null, 2), "utf8");

const event = {
  timestamp: now,
  action,
  status: nextState.status,
  activeBatch: nextState.activeBatch,
  deadline: nextState.deadline,
  lastCompletedCheckpoint: nextState.lastCompletedCheckpoint,
  nextIntendedAction: nextState.nextIntendedAction,
  note: nextState.note,
};

await appendFile(eventsPath, `${JSON.stringify(event)}\n`, "utf8");

console.log(
  `Updated ${path.relative(rootDir, statePath).replace(/\\/g, "/")} with ${action} (${status}).`,
);
