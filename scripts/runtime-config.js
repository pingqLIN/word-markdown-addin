import net from "node:net";
import path from "node:path";
import { mkdir, readFile, writeFile } from "node:fs/promises";

export const DEFAULT_LOCAL_PORT = 3000;
export const DEFAULT_LOCAL_HOST = `http://localhost:${DEFAULT_LOCAL_PORT}`;

const runtimeDirectoryName = ".local";
const runtimeConfigFilename = "runtime.json";

export const normalizeHost = (host) =>
  String(host || DEFAULT_LOCAL_HOST).trim().replace(/\/+$/, "");

export const getRuntimeConfigPath = (root = process.cwd()) =>
  path.join(root, runtimeDirectoryName, runtimeConfigFilename);

export const readRuntimeConfig = async (root = process.cwd()) => {
  try {
    const contents = await readFile(getRuntimeConfigPath(root), "utf8");
    const config = JSON.parse(contents);
    return config && typeof config === "object" ? config : {};
  } catch (error) {
    if (error?.code === "ENOENT") {
      return {};
    }

    throw error;
  }
};

export const readRuntimeHost = async (root = process.cwd()) => {
  const config = await readRuntimeConfig(root);
  if (typeof config.addinHost === "string" && config.addinHost.trim()) {
    return normalizeHost(config.addinHost);
  }

  return null;
};

export const writeRuntimeHost = async (host, root = process.cwd()) => {
  const normalizedHost = normalizeHost(host);
  const runtimeConfigPath = getRuntimeConfigPath(root);

  await mkdir(path.dirname(runtimeConfigPath), { recursive: true });
  await writeFile(
    runtimeConfigPath,
    JSON.stringify(
      {
        addinHost: normalizedHost,
        updatedAt: new Date().toISOString(),
      },
      null,
      2,
    ),
    "utf8",
  );

  return normalizedHost;
};

export const buildLocalHost = (port) => `http://localhost:${port}`;

export const buildProbeUrls = (host, routePath) => {
  const normalizedHost = normalizeHost(host);
  const normalizedPath = routePath.startsWith("/") ? routePath : `/${routePath}`;
  const baseUrl = new URL(normalizedHost);
  const urls = [`${normalizedHost}${normalizedPath}`];

  if (baseUrl.protocol === "http:" && baseUrl.hostname === "localhost") {
    urls.push(`http://127.0.0.1:${baseUrl.port}${normalizedPath}`);
    urls.push(`http://[::1]:${baseUrl.port}${normalizedPath}`);
  }

  return [...new Set(urls)];
};

export const isLocalHttpHost = (host) => {
  const parsed = new URL(normalizeHost(host));
  return parsed.protocol === "http:" && (
    parsed.hostname === "localhost" ||
    parsed.hostname === "127.0.0.1" ||
    parsed.hostname === "::1" ||
    parsed.hostname === "[::1]"
  );
};

export const getPortFromHost = (host) => {
  const parsed = new URL(normalizeHost(host));
  if (parsed.port) {
    return Number(parsed.port);
  }

  return parsed.protocol === "https:" ? 443 : 80;
};

export const isPortInUse = (port, host = "127.0.0.1") =>
  new Promise((resolvePort) => {
    const socket = net.createConnection({ port, host });

    socket.once("connect", () => {
      socket.destroy();
      resolvePort(true);
    });

    socket.once("error", () => {
      resolvePort(false);
    });
  });

export const isAnyLocalPortInUse = async (port) => {
  for (const host of ["127.0.0.1", "localhost", "::1"]) {
    if (await isPortInUse(port, host)) {
      return true;
    }
  }

  return false;
};

export const findAvailableLocalHost = async (
  preferredPort = DEFAULT_LOCAL_PORT,
  maxAttempts = 25,
) => {
  for (let offset = 0; offset < maxAttempts; offset += 1) {
    const port = preferredPort + offset;
    if (!await isAnyLocalPortInUse(port)) {
      return buildLocalHost(port);
    }
  }

  throw new Error(`No free localhost port found starting at ${preferredPort}.`);
};
