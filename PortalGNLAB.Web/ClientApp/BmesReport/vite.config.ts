import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

export default defineConfig({
  plugins: [react()],
  define: {
    "process.env.NODE_ENV": JSON.stringify("production"),
  },
  build: {
    outDir: resolve(import.meta.dirname, "../../wwwroot/bmes-report"),
    emptyOutDir: true,
    sourcemap: false,
    minify: "oxc",
    lib: {
      entry: resolve(import.meta.dirname, "src/main.tsx"),
      formats: ["es"],
      fileName: "bmes-report",
      cssFileName: "bmes-report",
    },
    rolldownOptions: {
      output: {
        codeSplitting: false,
      },
    },
  },
});
