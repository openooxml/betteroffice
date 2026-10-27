import { describe, expect, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, normalize } from "node:path";

const repoRoot = join(import.meta.dir, "..", "..", "..");
const dockerfilePath = join(import.meta.dir, "..", "Dockerfile");

function readDockerfile(): string {
  return readFileSync(dockerfilePath, "utf8");
}

function stages(source: string): string[][] {
  const result: string[][] = [];
  let current: string[] | null = null;
  for (const line of source.split("\n")) {
    if (/^FROM\s/i.test(line)) {
      current = [line];
      result.push(current);
    } else if (current) {
      current.push(line);
    }
  }
  return result;
}

interface CopyEntry {
  source: string;
  dest: string;
}

function copies(stageLines: string[]): CopyEntry[] {
  const entries: CopyEntry[] = [];
  for (const raw of stageLines) {
    const line = raw.trim();
    if (!/^COPY\s/i.test(line)) continue;
    const tokens = line.replace(/^COPY\s+/i, "").split(/\s+/).filter((token) => token.length > 0);
    const positional = tokens.filter((token) => !token.startsWith("--"));
    const dest = positional[positional.length - 1]!;
    for (const source of positional.slice(0, -1)) entries.push({ source, dest });
  }
  return entries;
}

const RUNTIME_ALLOWED = new Set([
  "/app/apps/wopi/package.json",
  "/app/apps/wopi/src",
  "/app/apps/wopi/web/dist",
  "/app/packages/docx/src",
  "/app/packages/fonts/assets",
  "/app/packages/fonts-cjk/assets",
  "/app/packages/pptx/src",
  "/app/packages/xlsx/src",
  "/app/shared/collaboration-limits.ts",
  "/app/shared/collaboration-pipe",
  "/app/shared/media.ts",
  "/app/shared/office-session",
  "/manifests",
]);

function runtimeCoverage(runtime: string[][]): string[] {
  const prefixes: string[] = [];
  for (const entry of copies(runtime)) {
    if (entry.source === "/manifests") continue;
    if (!entry.source.startsWith("/app/")) continue;
    prefixes.push(entry.source.slice("/app/".length));
  }
  return prefixes;
}

const SPECIFIER = /(?:import|export)\s+(type\s+)?[^"']*?from\s*["']([^"']+)["']|import\s*["']([^"']+)["']|import\s*\(\s*["']([^"']+)["']\)/g;

function resolveRelative(fromFile: string, specifier: string): string {
  const base = normalize(join(dirname(fromFile), specifier));
  const candidates = [
    base,
    `${base}.ts`,
    `${base}.tsx`,
    `${base}.js`,
    `${base}.mjs`,
    join(base, "index.ts"),
    join(base, "index.tsx"),
    join(base, "index.js"),
  ];
  for (const candidate of candidates) {
    if (existsSync(join(repoRoot, candidate))) return candidate.split("\\").join("/");
  }
  return base.split("\\").join("/");
}

function walkGraph(): { files: Set<string>; bare: Set<string> } {
  const files = new Set<string>();
  const bare = new Set<string>();
  const stack = ["apps/wopi/src/server.ts"];
  while (stack.length > 0) {
    const file = stack.pop()!;
    if (files.has(file)) continue;
    files.add(file);
    const absolute = join(repoRoot, file);
    if (!existsSync(absolute)) continue;
    const text = readFileSync(absolute, "utf8");
    for (const match of text.matchAll(SPECIFIER)) {
      const isType = Boolean(match[1]);
      const specifier = match[2] ?? match[3] ?? match[4];
      if (!specifier || isType) continue;
      if (specifier.startsWith("node:") || specifier === "bun:test") continue;
      if (specifier.startsWith("./") || specifier.startsWith("../")) {
        stack.push(resolveRelative(file, specifier));
      } else {
        const name = specifier.startsWith("@")
          ? specifier.split("/").slice(0, 2).join("/")
          : specifier.split("/")[0]!;
        bare.add(name);
      }
    }
  }
  return { files, bare };
}

describe("apps/wopi/Dockerfile", () => {
  test("is multi-stage and pins every base image by digest", () => {
    const source = readDockerfile();
    const fromLines = source.split("\n").filter((line) => /^FROM\s/i.test(line));
    expect(fromLines.length).toBeGreaterThanOrEqual(2);
    for (const line of fromLines) {
      expect(line).toMatch(/@sha256:[0-9a-f]{64}(\s|$)/);
    }
  });

  test("build stage is platform pinned and runtime installs production deps per target", () => {
    const all = stages(readDockerfile());
    expect(all[0]![0]).toContain("--platform=$BUILDPLATFORM");
    const runtimeText = all[all.length - 1]!.join("\n");
    expect(runtimeText).not.toContain("$BUILDPLATFORM");
    expect(runtimeText).not.toContain("x86_64");
    expect(runtimeText).toContain("bun install");
    expect(runtimeText).toContain("--production");
    expect(runtimeText).toContain("--frozen-lockfile");
  });

  test("runtime copies only allowed paths and the import graph stays inside them", () => {
    const all = stages(readDockerfile());
    const runtime = all[all.length - 1]!;
    const runtimeText = runtime.join("\n");
    expect(runtimeText).not.toMatch(/^COPY\s+(?!--from)/im);
    const entries = copies(runtime);
    expect(entries.length).toBeGreaterThan(0);
    for (const entry of entries) {
      expect(RUNTIME_ALLOWED.has(entry.source)).toBe(true);
    }
    const prefixes = runtimeCoverage(runtime);
    const { files, bare } = walkGraph();
    expect(files.size).toBeGreaterThan(50);
    expect(files.has("shared/office-session/types.ts")).toBe(true);
    expect(files.has("shared/media.ts")).toBe(true);
    for (const file of files) {
      const covered = prefixes.some((prefix) => file === prefix || file.startsWith(`${prefix}/`));
      expect(covered).toBe(true);
    }
    const pkg = JSON.parse(readFileSync(join(repoRoot, "apps/wopi/package.json"), "utf8")) as {
      dependencies: Record<string, string>;
    };
    for (const name of bare) {
      expect(Object.hasOwn(pkg.dependencies, name)).toBe(true);
    }
    expect(bare.has("jszip")).toBe(true);
  });

  test("test files are pruned before the runtime copy", () => {
    const source = readDockerfile();
    expect(source).toContain("*.test.");
    expect(source).toMatch(/find[\s\S]*-delete/);
  });

  test("build stage uses the pinned toolchains and builds wasm plus web", () => {
    const source = readDockerfile();
    expect(source).toContain("1.4.3");
    expect(source).toContain("wasm-pack");
    expect(source).toContain("0.15.0");
    expect(source).toContain("version_132");
    expect(source).toContain("195ddc94f9bc89f45abdabb0b9eea86023d727ba90eac8b35b80f2544fc30572");
    expect(source).toContain("c09f971ecaed9a2efc80fdcea7a00ef6b53c7fadc8c57d1f61b53a6aa66b668a");
    expect(source.match(/sha256sum -c/g)?.length).toBeGreaterThanOrEqual(2);
    expect(source).toMatch(/wasm-pack[\s\S]*--retry/);
    expect(source).toContain("rm -rf /tmp/wasm-pack");
    expect(source).toContain("bun install --frozen-lockfile");
    expect(source).toContain("build:xlsx-wasm");
    expect(source).toContain("build:docx-wasm");
    expect(source).toContain("build:pptx-wasm");
    expect(source).toContain("build:packages");
    expect(source).toContain("build:web");
    expect(source.indexOf("build:pptx-wasm")).toBeLessThan(source.indexOf("build:packages"));
    expect(source.indexOf("build:packages")).toBeLessThan(source.indexOf("build:web"));
    expect(source).toContain("wasm32-unknown-unknown");
  });

  test("runtime stage is non-root with port, volume, proof key default and healthcheck", () => {
    const source = readDockerfile();
    const all = stages(source);
    const runtime = all[all.length - 1]!.join("\n");
    expect(runtime).toMatch(/^EXPOSE 3000$/m);
    expect(runtime).toMatch(/^VOLUME \[?["']?\/data/m);
    expect(runtime).toContain("WOPI_PROOF_KEY_FILE=/data/proof-key.pem");
    expect(runtime).toMatch(/^HEALTHCHECK[\s\S]*\/healthz/m);
    expect(runtime).toMatch(/^USER (?!root\b|0\b)\S+$/m);
    expect(runtime).toContain("apps/wopi/src/server.ts");
    expect(runtime.toLowerCase()).not.toContain("access_token=");
  });

  test("dockerignore excludes generated wasm, secrets and fixture trees", () => {
    const lines = readFileSync(join(repoRoot, ".dockerignore"), "utf8")
      .split("\n")
      .map((line) => line.trim());
    for (const entry of [
      "packages/*/src/wasm/generated/**",
      "**/.dev.vars",
      "**/.next",
      "**/.open-next",
      "**/.wrangler",
      "**/.cargo-target",
      "fuzz",
      "crates/*/tests/fixtures/**",
    ]) {
      expect(lines).toContain(entry);
    }
  });

  test("no em or en dashes", () => {
    const source = readDockerfile();
    expect(source.includes("\u2014")).toBe(false);
    expect(source.includes("\u2013")).toBe(false);
  });
});
