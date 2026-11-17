// Copy the built add-in into the web app so townshipcanada.com serves it at
// /integrations/excel/ (the URLs in manifest.xml). Run `npm run publish:web`,
// then commit public/integrations/excel/ in the web repo and merge it to main.
// WEB_DIR overrides the web checkout location (default: ../web).
import { cpSync, existsSync, rmSync } from "node:fs";
import { resolve } from "node:path";

const webDir = resolve(process.env.WEB_DIR || "../web");
const target = resolve(webDir, "public/integrations/excel");

if (!existsSync(resolve(webDir, "nuxt.config.ts"))) {
  console.error(`No web checkout at ${webDir}. Set WEB_DIR to the townshipcanada/web clone.`);
  process.exit(1);
}
if (!existsSync("dist/manifest.xml")) {
  console.error("dist/ is missing. Run npm run build first.");
  process.exit(1);
}
rmSync(target, { recursive: true, force: true });
cpSync("dist", target, { recursive: true });
console.log(`Copied dist/ to ${target}`);
