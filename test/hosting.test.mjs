// Checks that the built add-in matches what manifest.xml tells Office to load.
// The add-in is hosted at https://townshipcanada.com/integrations/excel/ (the
// web app's public/integrations/excel/), so every manifest URL under that path
// must exist in dist/. Run after `npm run build`.
import { test } from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

const HOST = "https://townshipcanada.com/integrations/excel/";
const manifest = readFileSync("manifest.xml", "utf8");
const dist = "dist";

test("dist exists (run npm run build first)", () => {
  assert.ok(existsSync(dist), "dist/ is missing");
});

test("every add-in URL in the manifest is a file in dist/", () => {
  const urls = [...new Set(manifest.match(/https:\/\/townshipcanada\.com\/integrations\/excel\/[^"<\s]+/g))];
  assert.ok(urls.length >= 8, `expected the add-in URLs, found ${urls.length}`);
  const missing = urls.filter((u) => !existsSync(join(dist, u.slice(HOST.length))));
  assert.deepEqual(missing, []);
});

test("the manifest's other townshipcanada.com links are real pages, not the add-in path", () => {
  const others = (manifest.match(/https:\/\/townshipcanada\.com[^"<\s]*/g) || []).filter(
    (u) => !u.startsWith(HOST)
  );
  assert.ok(!others.includes("https://townshipcanada.com/integrations/excel"));
  assert.ok(!others.includes("https://townshipcanada.com/support"));
});

test("custom functions load in the shared runtime page", () => {
  const page = manifest.match(/<ExtensionPoint xsi:type="CustomFunctions">[\s\S]*?<Page>\s*<SourceLocation resid="([^"]+)"/);
  assert.equal(page?.[1], "Taskpane.Url");
  const runtime = manifest.match(/<Runtime resid="([^"]+)"/);
  assert.equal(runtime?.[1], "Taskpane.Url");
});

test("built pages load their scripts and styles relative to the hosting path", () => {
  for (const page of ["taskpane.html", "commands.html"]) {
    const html = readFileSync(join(dist, page), "utf8");
    assert.doesNotMatch(html, /(src|href)="\/(assets|src)\//, `${page} uses root-relative asset paths`);
  }
});

test("functions.js is a standalone script that registers the custom functions", () => {
  const js = readFileSync(join(dist, "functions.js"), "utf8");
  assert.doesNotMatch(js, /^\s*import\s/m, "functions.js must not use ES module imports");
  assert.match(js, /CustomFunctions\.associate\(\s*["']CONVERT["']/);
});
