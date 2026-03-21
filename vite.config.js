import { defineConfig } from "vite";
import { resolve } from "path";
import { viteStaticCopy } from "vite-plugin-static-copy";

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
        functions: resolve(__dirname, "functions.html"),
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
