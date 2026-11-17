import { defineConfig } from "vite";
import { resolve } from "path";

// Standalone functions.js for the manifest's Functions.Script.Url. Office loads
// it as a classic script, so it is one self-contained IIFE with no imports.
// Runs after the main build and adds to dist/ without clearing it.
export default defineConfig({
  publicDir: false,
  build: {
    outDir: "dist",
    emptyOutDir: false,
    lib: {
      entry: resolve(__dirname, "src/functions/functions.js"),
      name: "TownshipCanadaFunctions",
      formats: ["iife"],
      fileName: () => "functions.js"
    }
  }
});
