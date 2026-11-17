import { defineConfig } from "vite";
import { resolve } from "path";
import { viteStaticCopy } from "vite-plugin-static-copy";

// The add-in is served from https://townshipcanada.com/integrations/excel/
// (the web app's public/integrations/excel/, see scripts/publish-to-web.mjs),
// so built pages reference their assets relative to that path.
export default defineConfig({
  base: "./",
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
        commands: resolve(__dirname, "commands.html")
      }
    },
    outDir: "dist"
  },
  server: {
    port: 3000,
    https: true,
    headers: {
      "Access-Control-Allow-Origin": "*"
    }
  },
  plugins: [
    viteStaticCopy({
      targets: [
        { src: "src/functions/functions.json", dest: "." },
        { src: "manifest.xml", dest: "." }
      ]
    })
  ]
});
