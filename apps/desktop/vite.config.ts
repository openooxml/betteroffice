import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  build: { target: "safari16.4" },
  worker: { format: "es" },
});
