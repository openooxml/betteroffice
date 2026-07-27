import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

const ROOT = join(import.meta.dir, "..", "..", "..");
const LLMS = readFileSync(join(ROOT, "apps/web/public/llms.txt"), "utf8");

function publishedPackages() {
  const dir = join(ROOT, "packages");
  const names = [];
  for (const entry of readdirSync(dir)) {
    let manifest;
    try {
      manifest = JSON.parse(readFileSync(join(dir, entry, "package.json"), "utf8"));
    } catch {
      continue;
    }
    if (manifest.private) continue;
    if (manifest.name?.startsWith("@betteroffice/")) names.push(manifest.name);
  }
  return names.sort();
}

function publishedCrates() {
  const dir = join(ROOT, "crates");
  const names = [];
  for (const entry of readdirSync(dir)) {
    let manifest;
    try {
      manifest = readFileSync(join(dir, entry, "Cargo.toml"), "utf8");
    } catch {
      continue;
    }
    if (!manifest.includes('publish = ["crates-io"]')) continue;
    const name = manifest.match(/^name = "([^"]+)"/m)?.[1];
    if (name) names.push(name);
  }
  return names.sort();
}

describe("llms.txt", () => {
  test("names no npm package that does not exist", () => {
    const real = new Set(publishedPackages());
    const claimed = [...LLMS.matchAll(/@betteroffice\/[a-z0-9-]+/g)].map((m) => m[0]);
    const phantom = [...new Set(claimed)].filter((name) => !real.has(name));
    expect(phantom).toEqual([]);
  });

  test("names every published npm package", () => {
    const missing = publishedPackages().filter((name) => !LLMS.includes(name));
    expect(missing).toEqual([]);
  });

  test("names no crate that is not published", () => {
    const real = new Set(publishedCrates());
    const claimed = [...LLMS.matchAll(/\bbetteroffice(?:-[a-z0-9]+)+\b/g)].map((m) => m[0]);
    const phantom = [...new Set(claimed)].filter((name) => !real.has(name));
    expect(phantom).toEqual([]);
  });

  test("names every published crate", () => {
    const missing = publishedCrates().filter((name) => !LLMS.includes(name));
    expect(missing).toEqual([]);
  });
});
