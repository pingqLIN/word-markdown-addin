import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const manifestTemplatePath = path.resolve(process.cwd(), "manifest.template.xml");
const manifestOutputPath = path.resolve(process.cwd(), "manifest.xml");
const defaultHost = "http://localhost:3000";
const addinHost = (process.env.MANIFEST_HOST || defaultHost).replace(/\/+$/, "");

const templateContents = await readFile(manifestTemplatePath, "utf8");
const manifestContents = templateContents.replaceAll("{{ADDIN_HOST}}", addinHost);

await writeFile(manifestOutputPath, manifestContents, "utf8");
console.log(`Generated manifest.xml using ADDIN_HOST=${addinHost}`);
