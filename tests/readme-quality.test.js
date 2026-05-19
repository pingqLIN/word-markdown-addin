import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";

const repoRelativeImagePattern =
  /(?:!\[[^\]]*\]\((?!https?:\/\/)([^)]+)\)|<img[^>]+src="(?!https?:\/\/)([^"]+)")/g;

const assertFileExists = async (filePath) => {
  await access(filePath);
};

const extractRepoImagePaths = (markdown) =>
  [...markdown.matchAll(repoRelativeImagePattern)]
    .map((match) => match[1] || match[2])
    .filter(Boolean)
    .map((imagePath) => imagePath.replace(/^\.\/+/, ""));

test("README files follow the project documentation guardrails", async () => {
  const readmeFiles = ["README.md", "README.zh-TW.md"];

  for (const readmeFile of readmeFiles) {
    const contents = await readFile(readmeFile, "utf8");

    assert.match(contents, /^\[!\[Word Markdown Companion product banner\]/u);
    assert.match(contents, /img\.shields\.io\/badge\/version-0\.1\.0/u);
    assert.match(contents, /img\.shields\.io\/badge\/license-MIT/u);
    assert.match(contents, /## AI-Assisted Development/u);
    assert.match(contents, /\[MIT License\]\(LICENSE\)/u);
    assert.doesNotMatch(contents, /^This project is/mu);
    assert.doesNotMatch(contents, /assets\/mascot|word-markdown-companion-url-hero/u);

    for (const imagePath of extractRepoImagePaths(contents)) {
      await assertFileExists(path.resolve(imagePath));
    }
  }

  await assertFileExists("LICENSE");
});

test("README language switch links are reciprocal", async () => {
  const english = await readFile("README.md", "utf8");
  const traditionalChinese = await readFile("README.zh-TW.md", "utf8");

  assert.match(english, /\[繁體中文\]\(README\.zh-TW\.md\)/u);
  assert.match(traditionalChinese, /\[English\]\(README\.md\)/u);
});
