import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  define: {
    "globalThis.__DOCX_PEERLESS_HARNESS__": JSON.stringify(process.env.DOCX_PEERLESS_HARNESS === "1"),
  },
  build: { target: "safari16.4" },
  worker: { format: "es" },
});
