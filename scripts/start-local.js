import { spawn } from "node:child_process";

const NPM_COMMAND = process.platform === "win32" ? "npm.cmd" : "npm";
const defaultHost = "http://localhost:3000";
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

const runCommand = (command, cmdArgs, env) => new Promise((resolve, reject) => {
  const proc = spawn(command, cmdArgs, {
    env,
    stdio: "inherit",
    shell: false,
  });

  proc.on("error", (error) => reject(error));
  proc.on("close", (code) => {
    if (code === 0) {
      resolve(code);
      return;
    }
    reject(new Error(`command failed: ${command} ${cmdArgs.join(" ")} (exit code: ${code})`));
  });
});

const normalizedHost = addinHost.replace(/\/+$/, "");
const manifestEnv = { ...process.env, MANIFEST_HOST: normalizedHost };

console.log(`Word Markdown Add-in local start`);
console.log(`- ADDIN_HOST: ${normalizedHost}`);
console.log("Generating manifest.xml and starting server...");

try {
  await runCommand(NPM_COMMAND, ["run", "render-manifest"], manifestEnv);
  await runCommand(NPM_COMMAND, ["run", "dev-server"], manifestEnv);
} catch (error) {
  console.error(error?.message || error);
  process.exit(1);
}
