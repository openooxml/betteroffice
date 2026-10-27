import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  define: {
    "globalThis.__DOCX_PEERLESS_HARNESS__": JSON.stringify(process.env.DOCX_PEERLESS_HARNESS === "1"),
  },
  resolve: {
    alias: [{
      find: /^@betteroffice\/docx-react\/peerless-harness$/,
      replacement: fileURLToPath(new URL("../../packages/docx-react/dist/peerless-harness.mjs", import.meta.url)),
    }],
  },
  build: { target: "safari16.4" },
  worker: { format: "es" },
});
