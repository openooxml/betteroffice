import { describe, expect, test } from "bun:test";
import { mkdtemp, readdir, readFile, realpath, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { build } from "vite";
import desktopConfig from "../vite.config";
import wopiConfig from "../../wopi/web/vite.config";

describe("vite worker bundling", () => {
  test("desktop and wopi disable worker code splitting", () => {
    expect(desktopConfig.worker).toMatchObject({
      format: "es",
      rolldownOptions: { output: { codeSplitting: false } },
    });
    expect(wopiConfig.worker).toMatchObject({
      format: "es",
      rolldownOptions: { output: { codeSplitting: false } },
    });
  });

  test("worker build inlines dynamic imports instead of emitting back-importing chunks", async () => {
    const dir = await realpath(await mkdtemp(join(tmpdir(), "betteroffice-worker-bundle-")));
    try {
      await writeFile(join(dir, "index.html"), '<script type="module" src="./main.js"></script>');
      await writeFile(
        join(dir, "main.js"),
        'new Worker(new URL("./entry-worker.js", import.meta.url), { type: "module" });'
      );
      await writeFile(
        join(dir, "shared.js"),
        "export const state = { value: 1 }; export function bump() { return ++state.value; }"
      );
      await writeFile(
        join(dir, "lazy.js"),
        'import { bump } from "./shared.js"; export function runLazy() { return bump(); }'
      );
      await writeFile(
        join(dir, "entry-worker.js"),
        'import { state } from "./shared.js"; self.onmessage = async () => { const { runLazy } = await import("./lazy.js"); self.postMessage(state.value + runLazy()); };'
      );

      const outDir = join(dir, "dist");
      await build({
        configFile: false,
        root: dir,
        logLevel: "silent",
        build: { outDir, emptyOutDir: true },
        worker: desktopConfig.worker,
      });

      const assets = await readdir(join(outDir, "assets"));
      const workerFiles = assets.filter((name) => name.startsWith("entry-worker-") && name.endsWith(".js"));
      expect(workerFiles).toHaveLength(1);

      const jsAssets = assets.filter((name) => name.endsWith(".js"));
      for (const file of jsAssets) {
        if (file === workerFiles[0] || file.startsWith("index-")) continue;
        const source = await readFile(join(outDir, "assets", file), "utf8");
        expect(source).not.toContain(workerFiles[0]!);
      }

      const workerSource = await readFile(join(outDir, "assets", workerFiles[0]!), "utf8");
      expect(workerSource).toContain("runLazy");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
