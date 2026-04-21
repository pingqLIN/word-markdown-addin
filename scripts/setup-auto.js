import { spawn } from "node:child_process";
import { resolve } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { requestStatus } from "./http-probe.js";
import {
  buildProbeUrls,
  DEFAULT_LOCAL_HOST,
  findAvailableLocalHost,
  isLocalHttpHost,
  normalizeHost,
  readRuntimeHost,
  writeRuntimeHost,
} from "./runtime-config.js";

const NPM_COMMAND = process.platform === "win32" ? "npm.cmd" : "npm";
const args = process.argv.slice(2);

let addinHost = process.env.MANIFEST_HOST || DEFAULT_LOCAL_HOST;
let enableMdAssociation = false;
let hostWasExplicitlyProvided = Boolean(process.env.MANIFEST_HOST);

for (let index = 0; index < args.length; index += 1) {
  const token = args[index];

  if (token === "--host" && index + 1 < args.length) {
    addinHost = args[index + 1];
    hostWasExplicitlyProvided = true;
    index += 1;
    continue;
  }

  if (token.startsWith("--host=")) {
    addinHost = token.slice("--host=".length);
    hostWasExplicitlyProvided = true;
    continue;
  }

  if (token === "--associate-md" || token === "--with-md-association") {
    enableMdAssociation = true;
  }
}

const buildServerProbeUrls = (host) => buildProbeUrls(host, "/taskpane.html");
const buildServerCompatibilityUrls = (host) =>
  buildProbeUrls(host, "/api/pending-markdown");

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
  const response = await requestStatus(url);
  return response.ok;
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
    const response = await requestStatus(url);
    if (response.status === 200 || response.status === 204) {
      return true;
    }
  }

  return false;
};

const waitForUrl = async (urls, attempts = 30, intervalMs = 1000) => {
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (await isAnyUrlReady(urls)) {
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

const resolveLocalHost = async () => {
  const storedHost = await readRuntimeHost();
  const candidateHosts = [
    normalizeHost(addinHost),
    storedHost,
    DEFAULT_LOCAL_HOST,
  ].filter(Boolean);

  for (const candidateHost of [...new Set(candidateHosts)]) {
    if (!isLocalHttpHost(candidateHost)) {
      continue;
    }

    const probeUrls = buildServerProbeUrls(candidateHost);
    const compatibilityUrls = buildServerCompatibilityUrls(candidateHost);

    if (await isAnyUrlReady(probeUrls) && await isCompatibleServerReady(compatibilityUrls)) {
      await writeRuntimeHost(candidateHost);
      return candidateHost;
    }
  }

  const nextHost = await findAvailableLocalHost();
  await writeRuntimeHost(nextHost);
  return nextHost;
};

const main = async () => {
  let normalizedHost = normalizeHost(addinHost);

  if (!hostWasExplicitlyProvided && isLocalHttpHost(normalizedHost)) {
    normalizedHost = await resolveLocalHost();
  } else if (isLocalHttpHost(normalizedHost)) {
    await writeRuntimeHost(normalizedHost);
  }

  const manifestEnv = { ...process.env, MANIFEST_HOST: normalizedHost };
  const serverProbeUrls = buildServerProbeUrls(normalizedHost);
  const serverCompatibilityUrls = buildServerCompatibilityUrls(normalizedHost);

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

  if (isLocalHttpHost(normalizedHost) && await isAnyUrlReady(serverProbeUrls)) {
    if (await isCompatibleServerReady(serverCompatibilityUrls)) {
      console.log(`Detected an existing local dev server on ${normalizedHost}.`);
      console.log("Setup complete. Reusing the running server.");
      shouldReuseExistingServer = true;
    }
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
    throw error;
  }
};

try {
  await main();
} catch (error) {
  console.error(error?.message || error);
  process.exitCode = 1;
}
