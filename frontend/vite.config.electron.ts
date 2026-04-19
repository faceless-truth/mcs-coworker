/**
 * Vite config for building the Electron production bundle.
 * Key difference from the web build: base must be "./" (relative paths)
 * so assets load correctly from Electron's file:// protocol.
 *
 * Usage: npx vite build --config vite.config.electron.ts
 */
import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import path from "node:path";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react(), tailwindcss()],
  base: "./",  // Critical: relative paths for file:// loading in Electron
  resolve: {
    alias: {
      "@": path.resolve(import.meta.dirname, "client", "src"),
      "@shared": path.resolve(import.meta.dirname, "shared"),
    },
  },
  root: path.resolve(import.meta.dirname, "client"),
  build: {
    outDir: path.resolve(import.meta.dirname, "dist-electron"),
    emptyOutDir: true,
  },
});
