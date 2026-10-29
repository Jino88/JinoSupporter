import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

// The React app is served by PortalGNLAB.Web under /app (PortalShellEndpoints).
export default defineConfig({
  base: "/app/",
  plugins: [react()],
  build: {
    outDir: resolve(import.meta.dirname, "../../wwwroot/app"),
    emptyOutDir: true,
    sourcemap: false,
  },
  server: {
    proxy: {
      "/api": "http://localhost:5050",
      "/auth": "http://localhost:5050",
      "/standalone": "http://localhost:5050",
    },
  },
  test: {
    globals: true,
    environment: "node",
  },
});
