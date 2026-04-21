import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { setTimeout as delay } from "node:timers/promises";

export const repoRoot = new URL("..", import.meta.url);

export const runNodeCommand = async (args, options = {}) =>
  new Promise((resolve, reject) => {
    const child = spawn(process.execPath, args, {
      cwd: new URL(".", repoRoot),
      env: {
        ...process.env,
        ...options.env,
      },
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }

      reject(
        new Error(
          `Command failed (${code}): node ${args.join(" ")}\n${stdout}${stderr}`.trim(),
        ),
      );
    });
  });

export const getAvailablePort = async () => {
  const server = await import("node:net").then(({ createServer }) => createServer());
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  assert.ok(address && typeof address === "object" && address.port, "Expected an ephemeral port.");
  const { port } = address;
  server.close();
  await once(server, "close");
  return port;
};

export const waitFor = async (predicate, { attempts = 50, intervalMs = 100 } = {}) => {
  let lastError;

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    try {
      const value = await predicate();
      if (value) {
        return value;
      }
    } catch (error) {
      lastError = error;
    }

    await delay(intervalMs);
  }

  if (lastError) {
    throw lastError;
  }

  throw new Error("Timed out while waiting for predicate.");
};
