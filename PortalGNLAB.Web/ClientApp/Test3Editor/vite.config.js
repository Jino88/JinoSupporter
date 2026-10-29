import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

export default defineConfig({
  plugins: [react()],
  define: {
    "process.env.NODE_ENV": JSON.stringify("production")
  },
  build: {
    outDir: resolve(import.meta.dirname, "../../wwwroot/test3-editor"),
    emptyOutDir: true,
    sourcemap: false,
    minify: "oxc",
    lib: {
      entry: resolve(import.meta.dirname, "src/main.jsx"),
      formats: ["es"],
      fileName: "test3-editor",
      cssFileName: "test3-editor"
    }
  }
});
