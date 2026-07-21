import { defineConfig } from "vite";

export default defineConfig({
  root: "viewer-src",
  base: "/viewer/",
  server: {
    port: 5173,
    strictPort: true,
  },
  build: {
    outDir: "../viewer",
    emptyOutDir: true,
    target: "esnext",
  },
});
