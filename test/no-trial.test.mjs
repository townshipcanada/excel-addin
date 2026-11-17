// Township Canada does not offer trial API keys. The add-in must not mention
// them, route to a trial endpoint, or claim a paid key is unlimited.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";

function files(dir) {
  return readdirSync(dir).flatMap((name) => {
    const p = join(dir, name);
    return statSync(p).isDirectory() ? files(p) : [p];
  });
}

const sources = [...files("src"), "taskpane.html", "commands.html"].filter((p) =>
  /\.(js|html|css)$/.test(p)
);

test("no trial keys anywhere in the add-in source", () => {
  const hits = sources.filter((p) => /trial/i.test(readFileSync(p, "utf8")));
  assert.deepEqual(hits, []);
});

test("no claim that a paid key is unlimited", () => {
  const hits = sources.filter((p) => /unlimited/i.test(readFileSync(p, "utf8")));
  assert.deepEqual(hits, []);
});

test("every request goes to the public API", async () => {
  const config = await import("../src/shared/config.js");
  assert.equal(config.getApiBaseUrl(), "https://developer.townshipcanada.com");
});
