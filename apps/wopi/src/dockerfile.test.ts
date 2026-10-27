import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";

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

function copySources(stageLines: string[]): string[] {
  const sources: string[] = [];
  for (const raw of stageLines) {
    const line = raw.trim();
    if (!/^COPY\s/i.test(line)) continue;
    const tokens = line.replace(/^COPY\s+/i, "").split(/\s+/).filter((token) => token.length > 0);
    const positional = tokens.filter((token) => !token.startsWith("--"));
    for (const source of positional.slice(0, -1)) sources.push(source);
  }
  return sources;
}

const RUNTIME_ALLOWED = new Set([
  "/app/apps/wopi/package.json",
  "/app/apps/wopi/src",
  "/app/apps/wopi/web/dist",
  "/app/node_modules",
  "/app/packages/docx/src",
  "/app/packages/fonts/assets",
  "/app/packages/fonts-cjk/assets",
  "/app/packages/pptx/src",
  "/app/packages/xlsx/src",
  "/app/shared/collaboration-limits.ts",
  "/app/shared/collaboration-pipe",
]);

describe("apps/wopi/Dockerfile", () => {
  test("is multi-stage and pins every base image by digest", () => {
    const source = readDockerfile();
    const fromLines = source.split("\n").filter((line) => /^FROM\s/i.test(line));
    expect(fromLines.length).toBeGreaterThanOrEqual(2);
    for (const line of fromLines) {
      expect(line).toMatch(/@sha256:[0-9a-f]{64}(\s|$)/);
    }
  });

  test("runtime stage copies only allowed paths from the build stage", () => {
    const all = stages(readDockerfile());
    const runtime = all[all.length - 1]!;
    const runtimeText = runtime.join("\n");
    expect(runtimeText).not.toMatch(/^COPY\s+(?!--from)/im);
    const sources = copySources(runtime);
    expect(sources.length).toBeGreaterThan(0);
    for (const source of sources) {
      expect(RUNTIME_ALLOWED.has(source)).toBe(true);
    }
    for (const required of RUNTIME_ALLOWED) {
      expect(sources).toContain(required);
    }
  });

  test("build stage uses the pinned toolchains and builds wasm plus web", () => {
    const source = readDockerfile();
    expect(source).toContain("1.4.3");
    expect(source).toContain("wasm-pack");
    expect(source).toContain("0.15.0");
    expect(source).toContain("version_132");
    expect(source).toContain("195ddc94f9bc89f45abdabb0b9eea86023d727ba90eac8b35b80f2544fc30572");
    expect(source).toContain("sha256sum -c");
    expect(source).toContain("bun install --frozen-lockfile");
    expect(source).toContain("build:xlsx-wasm");
    expect(source).toContain("build:docx-wasm");
    expect(source).toContain("build:pptx-wasm");
    expect(source).toContain("build:web");
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

  test("no em or en dashes", () => {
    const source = readDockerfile();
    expect(source.includes("\u2014")).toBe(false);
    expect(source.includes("\u2013")).toBe(false);
  });
});
