import { fileURLToPath } from "node:url";
import { defineConfig, type Plugin } from "vite";

const webRoot = fileURLToPath(new URL(".", import.meta.url));

/** Keeps bundled font files out of the bundle; they are served from /fonts/. */
function fontsExternalPlugin(): Plugin {
  return {
    name: "wopi-fonts-external",
    enforce: "pre",
    transform(code, id) {
      if (!id.includes("packages/fonts") && !id.includes("@betteroffice/fonts")) return null;
      if (!code.includes("assets/")) return null;
      const rewritten = code.replace(
        /new URL\((['"])(?:\.\.\/|\.\/)assets\//g,
        "new URL($1/fonts/",
      );
      return rewritten === code ? null : { code: rewritten, map: null };
    },
  };
}

export default defineConfig({
  root: webRoot,
  base: "/app/",
  plugins: [fontsExternalPlugin()],
  build: {
    target: "safari16.4",
    outDir: fileURLToPath(new URL("dist", import.meta.url)),
    emptyOutDir: true,
    cssCodeSplit: false,
    rollupOptions: {
      output: {
        entryFileNames: "main.js",
        chunkFileNames: "assets/[name]-[hash].js",
        assetFileNames: (info) =>
          info.names?.some((name) => name.endsWith(".css"))
            ? "main.css"
            : "assets/[name]-[hash][extname]",
      },
    },
  },
  worker: { format: "es" },
});
