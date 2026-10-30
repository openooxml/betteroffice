import { describe, expect, test } from "bun:test";
import { readFileSync, statSync } from "node:fs";
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

function isFile(repoRelative: string): boolean {
  try {
    return statSync(join(repoRoot, repoRelative)).isFile();
  } catch {
    return false;
  }
}

function runtimeCoverage(runtime: string[]): string[] {
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
    if (isFile(candidate)) return candidate.split("\\").join("/");
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
    if (!isFile(file)) continue;
    const text = readFileSync(join(repoRoot, file), "utf8");
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
    const manifests = copies(runtime).some((entry) => entry.source === "/manifests");
    for (const file of files) {
      const covered = prefixes.some((prefix) => file === prefix || file.startsWith(`${prefix}/`))
        || (manifests && /^(apps|packages|bindings|crates)\/.+\/package\.json$/.test(file));
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

  test("build tools select pinned downloads for each build architecture", () => {
    const build = stages(readDockerfile())[0]!.join("\n");
    expect(build).toMatch(/^ARG BUILDARCH$/m);
    for (const tool of [
      {
        name: "wasm-pack",
        variable: "wasm_pack",
        url: "https://github.com/rustwasm/wasm-pack/releases/download/v0.15.0/wasm-pack-v0.15.0-${wasm_pack_arch}-unknown-linux-musl.tar.gz",
        amd64: "c09f971ecaed9a2efc80fdcea7a00ef6b53c7fadc8c57d1f61b53a6aa66b668a",
        arm64: "e17ef0806381c3a0acb9c9ddad643a49facaa5a2ecf657a421d4d8f3357a24b7",
      },
      {
        name: "binaryen",
        variable: "binaryen",
        url: "https://github.com/WebAssembly/binaryen/releases/download/version_132/binaryen-version_132-${binaryen_arch}-linux.tar.gz",
        amd64: "195ddc94f9bc89f45abdabb0b9eea86023d727ba90eac8b35b80f2544fc30572",
        arm64: "c58562417836c5d0493d89bdefc434933bdc097db641b483df86bcfa557a107f",
      },
    ]) {
      const instruction = build.split(/^RUN /m).find((run) => run.includes(`-o /tmp/${tool.name}.tar.gz`))!;
      expect(instruction).toMatch(/^case "\$BUILDARCH" in \\/);
      expect(instruction).toContain(`amd64) ${tool.variable}_arch=x86_64; ${tool.variable}_sha256=${tool.amd64} ;;`);
      expect(instruction).toContain(`arm64) ${tool.variable}_arch=aarch64; ${tool.variable}_sha256=${tool.arm64} ;;`);
      expect(instruction).toContain('*) echo "Unsupported build architecture: $BUILDARCH" >&2; exit 1 ;;');
      expect(instruction).toContain(`"${tool.url}"`);
      expect(instruction).toContain(`echo "\${${tool.variable}_sha256}  /tmp/${tool.name}.tar.gz" | sha256sum -c -`);
    }
  });

  test("runtime code and assets are root owned and read only before switching to bun", () => {
    const all = stages(readDockerfile());
    const runtime = all[all.length - 1]!;
    for (const line of runtime.filter((line) => /^COPY\s/i.test(line))) {
      expect(line).toContain("--chown=root:root");
    }
    const text = runtime.join("\n").replace(/[ \t]*\\\n[ \t]*/g, " ");
    expect(text).not.toContain("--chown=bun");
    expect(text).toContain("chown -R root:root /app && chmod -R a-w /app");
    expect(text).toContain("mkdir -p /data && chown bun:bun /data");
    expect(text).toMatch(/^USER bun$/m);
    const permissions = text.indexOf("chmod -R a-w /app");
    expect(text.lastIndexOf("\nCOPY ")).toBeLessThan(permissions);
    expect(text.indexOf("bun install --production")).toBeLessThan(permissions);
    expect(permissions).toBeLessThan(text.indexOf("\nUSER bun\n"));
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
      "**/.env",
      "**/.env.*",
      "**/.dev.vars",
      "**/.next",
      "**/.open-next",
      "**/.wrangler",
      "**/.cargo-target",
      "fuzz",
      "data",
      "apps/wopi/data",
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
