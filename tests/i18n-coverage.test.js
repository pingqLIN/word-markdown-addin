import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const readJson = async (filePath) => JSON.parse(await readFile(filePath, "utf8"));

const readPath = (object, keyPath) =>
  keyPath
    .split(".")
    .reduce(
      (value, key) => (value && typeof value === "object" ? value[key] : undefined),
      object,
    );

test("taskpane i18n keys are covered by every locale", async () => {
  const taskpaneHtml = await readFile("src/taskpane.html", "utf8");
  const keys = [
    ...taskpaneHtml.matchAll(/data-i18n(?:-[a-z-]+)?="([^"]+)"/g),
  ].map((match) => match[1]);
  const uniqueKeys = [...new Set(keys)];
  const locales = {
    "zh-TW": await readJson("src/locales/zh-TW.json"),
    "en-US": await readJson("src/locales/en-US.json"),
  };

  for (const [locale, messages] of Object.entries(locales)) {
    const missingKeys = uniqueKeys.filter(
      (key) => typeof readPath(messages, key) !== "string",
    );

    assert.deepEqual(missingKeys, [], `${locale} should define every taskpane i18n key.`);
  }
});

test("English locale does not contain untranslated CJK UI copy", async () => {
  const enUs = await readFile("src/locales/en-US.json", "utf8");
  assert.equal(
    /[\u3400-\u9fff]/u.test(enUs),
    false,
    "en-US locale should not contain Chinese UI copy.",
  );
});
