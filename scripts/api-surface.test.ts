import { describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { buildSnapshot, generateSnapshotsForPackage, reachableFiles } from "./api-surface";

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
});
