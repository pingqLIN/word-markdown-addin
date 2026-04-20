import { spawn } from "node:child_process";
import net from "node:net";
import { resolve } from "node:path";
import { setTimeout as delay } from "node:timers/promises";

const NPM_COMMAND = process.platform === "win32" ? "npm.cmd" : "npm";
const defaultHost = "http://localhost:3000";
const args = process.argv.slice(2);

let addinHost = process.env.MANIFEST_HOST || defaultHost;
let enableMdAssociation = false;

for (let index = 0; index < args.length; index += 1) {
  const token = args[index];

  if (token === "--host" && index + 1 < args.length) {
    addinHost = args[index + 1];
    index += 1;
    continue;
  }

  if (token.startsWith("--host=")) {
    addinHost = token.slice("--host=".length);
    continue;
  }

  if (token === "--associate-md" || token === "--with-md-association") {
    enableMdAssociation = true;
  }
}

const normalizedHost = addinHost.replace(/\/+$/, "");
const manifestEnv = { ...process.env, MANIFEST_HOST: normalizedHost };
const serverProbeUrls = [
  "http://localhost:3000/taskpane.html",
  "http://127.0.0.1:3000/taskpane.html",
  "http://[::1]:3000/taskpane.html",
];
const serverCompatibilityUrls = [
  "http://localhost:3000/api/pending-markdown",
  "http://127.0.0.1:3000/api/pending-markdown",
  "http://[::1]:3000/api/pending-markdown",
];

const quoteWindowsArgument = (value) =>
  /[\s"]/u.test(value) ? `"${value.replace(/"/gu, "\"\"")}"` : value;

const runCommand = (command, commandArgs, env = process.env) => new Promise((resolveCommand, reject) => {
  const proc = process.platform === "win32" && /\.cmd$/i.test(command)
    ? spawn("cmd.exe", [
      "/d",
      "/s",
      "/c",
      `${command} ${commandArgs.map(quoteWindowsArgument).join(" ")}`.trim(),
    ], {
      env,
      stdio: "inherit",
    })
    : spawn(command, commandArgs, {
      env,
      stdio: "inherit",
    });

  proc.on("error", (error) => reject(error));
  proc.on("close", (code) => {
    if (code === 0) {
      resolveCommand(code);
      return;
    }
    reject(new Error(`command failed: ${command} ${commandArgs.join(" ")} (exit code: ${code})`));
  });
});

const isUrlReady = async (url) => {
  try {
    const response = await fetch(url, { method: "GET" });
    return response.ok;
  } catch {
    return false;
  }
};

const isAnyUrlReady = async (urls) => {
  for (const url of urls) {
    if (await isUrlReady(url)) {
      return true;
    }
  }

  return false;
};

const isCompatibleServerReady = async (urls) => {
  for (const url of urls) {
    try {
      const response = await fetch(url, { method: "GET" });
      if (response.status === 200 || response.status === 204) {
        return true;
      }
    } catch {
      // Ignore individual probe failures and continue.
    }
  }

  return false;
};

const isPortInUse = (port, host) => new Promise((resolvePort) => {
  const socket = net.createConnection({ port, host });

  socket.once("connect", () => {
    socket.destroy();
    resolvePort(true);
  });

  socket.once("error", () => {
    resolvePort(false);
  });
});

const isLocalPortInUse = async (port) => {
  for (const host of ["127.0.0.1", "localhost", "::1"]) {
    if (await isPortInUse(port, host)) {
      return true;
    }
  }

  return false;
};

const waitForUrl = async (urls, attempts = 30, intervalMs = 1000) => {
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (await isAnyUrlReady(urls) && await isCompatibleServerReady(serverCompatibilityUrls)) {
      return;
    }

    await delay(intervalMs);
  }

  throw new Error(`dev server did not become ready: ${urls.join(", ")}`);
};

const startBackgroundProcess = (command, commandArgs, env = process.env) => {
  if (process.platform === "win32" && /\.cmd$/i.test(command)) {
    return spawn("cmd.exe", [
      "/d",
      "/s",
      "/c",
      `${command} ${commandArgs.map(quoteWindowsArgument).join(" ")}`.trim(),
    ], {
      env,
      stdio: "inherit",
    });
  }

  return spawn(command, commandArgs, {
    env,
    stdio: "inherit",
  });
};

const main = async () => {
  console.log("Word Markdown Add-in auto setup");
  console.log(`- ADDIN_HOST: ${normalizedHost}`);
  console.log("Running npm install...");
  await runCommand(NPM_COMMAND, ["install", "--no-audit", "--no-fund"]);

  console.log("Generating manifest.xml...");
  await runCommand(NPM_COMMAND, ["run", "render-manifest"], manifestEnv);

  if (enableMdAssociation) {
    if (process.platform === "win32") {
      const associationScript = resolve(process.cwd(), "scripts", "enable-md-association.ps1");
      await runCommand("powershell", [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        associationScript,
      ]);
    } else {
      console.log("非 Windows 環境無法自動設定 .md 檔關聯，將略過該步驟。");
    }
  }

  let shouldReuseExistingServer = false;

  if (normalizedHost === defaultHost && await isAnyUrlReady(serverProbeUrls)) {
    if (!await isCompatibleServerReady(serverCompatibilityUrls)) {
      throw new Error("An older dev server is already running on http://localhost:3000. Stop the existing node process and rerun setup so the updated Markdown bridge API is available.");
    }

    console.log("Detected an existing local dev server on http://localhost:3000.");
    console.log("Setup complete. Reusing the running server.");
    shouldReuseExistingServer = true;
  }

  if (shouldReuseExistingServer) {
    return;
  }

  console.log("Starting dev server...");
  const serverProcess = startBackgroundProcess(NPM_COMMAND, ["run", "dev-server"], manifestEnv);

  try {
    await waitForUrl(serverProbeUrls);
    console.log("Setup complete. Keep this terminal open while using the add-in.");
    await new Promise((resolveServer) => {
      serverProcess.on("close", () => resolveServer());
    });
  } catch (error) {
    if (!serverProcess.killed) {
      serverProcess.kill("SIGINT");
    }
    if (normalizedHost === defaultHost && await isLocalPortInUse(3000)) {
      throw new Error("port 3000 is already in use by another process, but it is not serving this add-in. Stop that process or rerun with --host http://localhost:<free-port>.");
    }
    throw error;
  }
};

try {
  await main();
} catch (error) {
  console.error(error?.message || error);
  process.exitCode = 1;
}
