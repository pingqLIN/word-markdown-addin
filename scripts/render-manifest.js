import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const defaultHost = "http://localhost:3000";
const defaultManifestId = "e6c1ec6a-3b55-4ed6-8a57-1d3de4f6b4d1";
const defaultProviderName = "Internal";
const defaultDisplayName = "Word Markdown Companion";
const defaultDescription =
  "Import and export Markdown files in Microsoft Word with one click.";

const args = process.argv.slice(2);

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

const hasFlag = (name) => args.includes(`--${name}`);

const templatePath = path.resolve(
  process.cwd(),
  readOption("template") || "manifest.template.xml",
);
const outputPath = path.resolve(
  process.cwd(),
  readOption("output") || "manifest.xml",
);
const addinHost = (readOption("host") || process.env.MANIFEST_HOST || defaultHost).replace(
  /\/+$/,
  "",
);
const requireHttps = hasFlag("require-https");
const supportUrl = (readOption("support-url") || process.env.SUPPORT_URL || `${addinHost}/`).replace(
  /\/+$/,
  "/",
);
const replacements = {
  "{{ADDIN_HOST}}": addinHost,
  "{{ADDIN_ID}}": readOption("id") || process.env.ADDIN_ID || defaultManifestId,
  "{{PROVIDER_NAME}}":
    readOption("provider") || process.env.PROVIDER_NAME || defaultProviderName,
  "{{DISPLAY_NAME}}":
    readOption("display-name") || process.env.DISPLAY_NAME || defaultDisplayName,
  "{{DESCRIPTION}}":
    readOption("description") || process.env.ADDIN_DESCRIPTION || defaultDescription,
  "{{SUPPORT_URL}}": supportUrl,
};

if (requireHttps && !/^https:\/\//iu.test(addinHost)) {
  throw new Error(
    `A store or online manifest requires HTTPS. Received MANIFEST_HOST=${addinHost}`,
  );
}

if (requireHttps && !/^https:\/\//iu.test(supportUrl)) {
  throw new Error(
    `A store or online manifest requires an HTTPS support URL. Received SUPPORT_URL=${supportUrl}`,
  );
}

const templateContents = await readFile(templatePath, "utf8");
const manifestContents = Object.entries(replacements).reduce(
  (contents, [placeholder, value]) => contents.replaceAll(placeholder, value),
  templateContents,
);

await mkdir(path.dirname(outputPath), { recursive: true });
await writeFile(outputPath, manifestContents, "utf8");

console.log(`Generated ${path.relative(process.cwd(), outputPath)} using ADDIN_HOST=${addinHost}`);
