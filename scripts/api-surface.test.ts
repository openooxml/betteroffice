import { describe, expect, test } from "bun:test";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import {
  buildSharedSnapshots,
  buildSnapshot,
  entriesFromExports,
  FORMATS,
  generateSnapshotsForPackage,
  reachableFiles,
  sharedFileName,
} from "./api-surface";

describe("api-surface", () => {
  test("includes transitive declaration inputs and resolves JS specifiers, indexes and package paths", () => {
    const root = mkdtempSync(join(tmpdir(), "api-surface-fixture-"));
    const config = {
      packageDir: "fixture",
      packageName: "@test/fixture",
      entries: { ".": "src/index.ts" },
    };
    const write = (file: string, content: string) => {
      const path = join(root, config.packageDir, file);
      mkdirSync(dirname(path), { recursive: true });
      writeFileSync(path, content);
    };
    try {
      write("package.json", JSON.stringify({ name: config.packageName, type: "module" }));
      write("tsconfig.json", JSON.stringify({
        compilerOptions: {
          target: "ES2020", module: "ESNext", moduleResolution: "Bundler", noEmit: true,
          paths: { "@inputs/*": ["./src/inputs/*"] },
        },
        include: ["src"],
      }));
      write("src/index.ts", [
        'import type { Input } from "./generated/input.js";',
        'export type PublicInput = Input;',
        'export type { Alias } from "@inputs/alias.js";',
        'export type { DirectoryInput } from "./directory";',
        'export type Result = import("./result.js").Result;',
      ].join("\n"));
      write("src/generated/input.d.ts", [
        'import "./side-effect.js";',
        'import type { Nested } from "../nested";',
        'export interface Input { value: string; nested: Nested; }',
      ].join("\n"));
      write("src/generated/side-effect.d.ts", "export {}; declare global { interface FixtureGlobal { enabled: boolean; } }");
      write("src/nested/index.d.ts", "export interface Nested { count: number; }");
      write("src/inputs/alias.d.ts", "export interface Alias { label: string; }");
      write("src/directory/package.json", JSON.stringify({ types: "entry.d.ts" }));
      write("src/directory/entry.d.ts", "export interface DirectoryInput { count: number; }");
      write("src/result.ts", "export interface Result { ok: boolean; }");
      write("src/unused.d.ts", "export interface Unused { ignored: true; }");
      const snapshot = generateSnapshotsForPackage(config, root).get("index.api.d.ts") as string;
      expect(snapshot).toContain("// ---- generated/input.d.ts ----");
      expect(snapshot).toContain("value: string");
      expect(snapshot).toContain("// ---- generated/side-effect.d.ts ----");
      expect(snapshot).toContain("// ---- nested/index.d.ts ----");
      expect(snapshot).toContain("// ---- inputs/alias.d.ts ----");
      expect(snapshot).toContain("// ---- directory/entry.d.ts ----");
      expect(snapshot).toContain("// ---- result.d.ts ----");
      expect(snapshot).not.toContain("Unused");
      expect(snapshot).not.toContain(root);
      expect(generateSnapshotsForPackage(config, root).get("index.api.d.ts")).toBe(snapshot);
      write("src/generated/input.d.ts", "export interface Input { value: number; }");
      const changed = generateSnapshotsForPackage(config, root).get("index.api.d.ts") as string;
      expect(changed).not.toBe(snapshot);
      expect(changed).toContain("value: number");
      expect(changed).not.toContain("// ---- nested/index.d.ts ----");
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  test("fails clearly on an unresolved local declaration import", () => {
    const files = new Map([["index.d.ts", 'export type { Missing } from "./missing.js";']]);
    expect(() => buildSnapshot("@test/fixture", ".", "index.d.ts", files))
      .toThrow('Unresolved local declaration import "./missing.js" in index.d.ts');
  });

  test("orders reachable files independently of map insertion order", () => {
    const files = new Map([
      ["index.d.ts", 'export type { Z } from "./z.js"; export type { A } from "./a.js";'],
      ["z.d.ts", "export interface Z {}"],
      ["a.d.ts", "export interface A {}"],
    ]);
    expect(reachableFiles("index.d.ts", files)).toEqual(["a.d.ts", "index.d.ts", "z.d.ts"]);
    expect(buildSnapshot("@test/fixture", ".", "index.d.ts", files))
      .toBe(buildSnapshot("@test/fixture", ".", "index.d.ts", new Map([...files].reverse())));
  });

  test("shared snapshots keep each declaration file once and index it from every entry that reaches it", () => {
    const files = new Map([
      ["index.d.ts", 'export type { Shared } from "./shared.js"; export type { Own } from "./own.js";'],
      ["worker.d.ts", 'export type { Shared } from "./shared.js";'],
      ["shared.d.ts", "export interface Shared { value: string; }\n//# sourceMappingURL=shared.d.ts.map"],
      ["own.d.ts", "export interface Own {}"],
      ["unused.d.ts", "export interface Unused {}"],
    ]);
    const snapshots = buildSharedSnapshots("@test/fixture", [[".", "index.d.ts"], ["./worker", "worker.d.ts"]], files);
    expect([...snapshots.keys()].sort()).toEqual([
      "files/index.api.d.ts", "files/own.api.d.ts", "files/shared.api.d.ts", "files/worker.api.d.ts",
      "index.api.d.ts", "worker.api.d.ts",
    ]);
    expect(snapshots.get("index.api.d.ts")).toContain("files/shared.api.d.ts\n");
    expect(snapshots.get("worker.api.d.ts")).toContain("files/shared.api.d.ts\n");
    expect(snapshots.get("worker.api.d.ts")).not.toContain("files/own.api.d.ts");
    const shared = snapshots.get("files/shared.api.d.ts") as string;
    expect(shared).toContain("export interface Shared { value: string; }");
    expect(shared).not.toContain("sourceMappingURL");
    expect([...snapshots.values()].filter((content) => content.includes("interface Shared"))).toHaveLength(1);
    expect(buildSharedSnapshots("@test/fixture", [[".", "index.d.ts"], ["./worker", "worker.d.ts"]], new Map([...files].reverse())))
      .toEqual(snapshots);
    expect(sharedFileName("a/b.d.mts")).toBe("files/a/b.api.d.ts");
  });

  test("docx checks every sub-path its package exports, from the source tsup builds it from", () => {
    const docx = join(import.meta.dir, "..", "packages/docx");
    const manifest = JSON.parse(readFileSync(join(docx, "package.json"), "utf8")) as { exports: Record<string, unknown> };
    const tsup = readFileSync(join(docx, "tsup.config.ts"), "utf8");
    expect(Object.keys(FORMATS.docx.entries).sort()).toEqual(Object.keys(manifest.exports).sort());
    for (const source of Object.values(FORMATS.docx.entries)) {
      expect(existsSync(join(docx, source))).toBe(true);
      const built = source.replace(/^src\//, "").replace(/\.ts$/, "");
      expect(tsup).toMatch(new RegExp(`['"]?${built.replace(/[/.]/g, "\\$&")}['"]?: '${source.replace(/[/.]/g, "\\$&")}'`));
    }
  });

  test("an export without a dist types file fails clearly", () => {
    const dir = mkdtempSync(join(tmpdir(), "api-surface-exports-"));
    try {
      writeFileSync(join(dir, "package.json"), JSON.stringify({ exports: { ".": "./dist/index.d.mts", "./x": { import: "./x.mjs" } } }));
      expect(() => entriesFromExports(dir)).toThrow("./x: no ./dist/<path>.d.mts types file to map to its source");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
