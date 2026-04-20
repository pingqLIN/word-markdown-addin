import { spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import net from "node:net";
import { resolve } from "node:path";
import { setTimeout as delay } from "node:timers/promises";

const NPM_COMMAND = process.platform === "win32" ? "npm.cmd" : "npm";
const defaultHost = "http://localhost:3000";
const developerRegistryKey = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";
const args = process.argv.slice(2);

let addinHost = process.env.MANIFEST_HOST || defaultHost;

for (let index = 0; index < args.length; index += 1) {
  const token = args[index];
  if (token === "--host" && index + 1 < args.length) {
    addinHost = args[index + 1];
    index += 1;
    continue;
  }
  if (token.startsWith("--host=")) {
    addinHost = token.slice("--host=".length);
  }
}

const normalizedHost = addinHost.replace(/\/+$/, "");
const manifestEnv = { ...process.env, MANIFEST_HOST: normalizedHost };
const manifestPath = resolve(process.cwd(), "manifest.xml");
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

const runCommand = (command, commandArgs, env = process.env) =>
  new Promise((resolveCommand, reject) => {
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

const captureCommand = (command, commandArgs, env = process.env) =>
  new Promise((resolveCommand, reject) => {
    const proc = process.platform === "win32" && /\.cmd$/i.test(command)
      ? spawn("cmd.exe", [
        "/d",
        "/s",
        "/c",
        `${command} ${commandArgs.map(quoteWindowsArgument).join(" ")}`.trim(),
      ], {
        env,
        stdio: ["ignore", "pipe", "pipe"],
      })
      : spawn(command, commandArgs, {
        env,
        stdio: ["ignore", "pipe", "pipe"],
      });

    let stdout = "";
    let stderr = "";

    proc.stdout?.on("data", (chunk) => {
      stdout += chunk.toString();
    });

    proc.stderr?.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    proc.on("error", (error) => reject(error));
    proc.on("close", (code) => {
      if (code === 0) {
        resolveCommand({ stdout, stderr });
        return;
      }

      reject(new Error(stderr.trim() || `command failed: ${command} ${commandArgs.join(" ")} (exit code: ${code})`));
    });
  });

const startProcess = (command, commandArgs, env = process.env) =>
  process.platform === "win32" && /\.cmd$/i.test(command)
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

const hasWebViewLoopbackExemption = async () => {
  try {
    const { stdout } = await captureCommand("CheckNetIsolation", [
      "LoopbackExempt",
      "-s",
    ]);

    return stdout.toLowerCase().includes("microsoft.win32webviewhost_cw5n1h2txyewy");
  } catch {
    return false;
  }
};

const readManifestId = async () => {
  const contents = await readFile(manifestPath, "utf8");
  const match = contents.match(/<Id>\s*([^<\s]+)\s*<\/Id>/i);

  if (!match) {
    throw new Error(`unable to read add-in ID from ${manifestPath}`);
  }

  return match[1];
};

let serverProcess;

const shutdownServer = () => {
  if (!serverProcess || serverProcess.killed) {
    return;
  }
  serverProcess.kill("SIGINT");
};

process.on("SIGINT", () => {
  shutdownServer();
  process.exit(0);
});

process.on("SIGTERM", () => {
  shutdownServer();
  process.exit(0);
});

try {
  if (process.platform !== "win32") {
    throw new Error("desktop sideload is currently implemented only for Windows Office.");
  }

  console.log("Word Markdown Add-in sideload");
  console.log(`- MANIFEST_HOST: ${normalizedHost}`);

  if (normalizedHost === defaultHost && !await hasWebViewLoopbackExemption()) {
    console.warn("Warning: Microsoft Edge WebView loopback exemption was not detected.");
    console.warn("If Word later shows 'We can't open this add-in from localhost', run this in an elevated terminal:");
    console.warn('  CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"');
  }

  console.log("Generating manifest.xml...");
  await runCommand(NPM_COMMAND, ["run", "render-manifest"], manifestEnv);

  if (normalizedHost === defaultHost) {
    if (await isAnyUrlReady(serverProbeUrls)) {
      if (!await isCompatibleServerReady(serverCompatibilityUrls)) {
        throw new Error("An older dev server is already running on http://localhost:3000. Stop the existing node process and rerun sideload so the updated Markdown bridge API is available.");
      }

      console.log("Detected an existing local dev server on http://localhost:3000.");
    } else {
      console.log("Starting local dev server...");
      serverProcess = startProcess(NPM_COMMAND, ["run", "dev-server"], manifestEnv);
      try {
        await waitForUrl(serverProbeUrls);
      } catch (error) {
        if (await isLocalPortInUse(3000)) {
          throw new Error("port 3000 is already in use by another process, but it is not serving this add-in. Stop that process or rerun with --host http://localhost:<free-port>.");
        }
        throw error;
      }
    }
  } else {
    console.log("MANIFEST_HOST is not localhost; skipping local dev server startup.");
  }

  const addinId = await readManifestId();

  console.log("Registering manifest in Office developer sideload registry...");
  console.log(`- Registry key: ${developerRegistryKey}`);
  console.log(`- Value name: ${addinId}`);
  await runCommand("reg", [
    "add",
    developerRegistryKey,
    "/v",
    addinId,
    "/t",
    "REG_SZ",
    "/d",
    manifestPath,
    "/f",
  ]);

  console.log("Sideload registration complete.");
  console.log("Close all Word windows, then reopen Word.");
  console.log("If the ribbon button does not appear immediately, use Home > Add-ins once to refresh the add-in activation.");

  if (serverProcess) {
    console.log("Keep this terminal open while using the add-in.");
    await new Promise((resolveServer) => {
      serverProcess.on("close", () => resolveServer());
    });
  }
} catch (error) {
  shutdownServer();
  console.error(error?.message || error);
  process.exit(1);
}
