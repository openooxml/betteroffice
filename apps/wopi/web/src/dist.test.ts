import { describe, expect, test } from "bun:test";
import { existsSync, readdirSync, statSync } from "node:fs";
import { join, resolve } from "node:path";

const DIST = resolve(import.meta.dir, "../dist");
const FONT_PATTERN = /\.(ttf|otf|woff|woff2)$/i;

function walk(dir: string, prefix = ""): Array<{ path: string; size: number }> {
  const out: Array<{ path: string; size: number }> = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) out.push(...walk(full, rel));
    else out.push({ path: rel, size: statSync(full).size });
  }
  return out;
}

describe("web dist", () => {
  test("contains no font files and lists the largest files", () => {
    if (!existsSync(DIST)) {
      console.log("web dist not built, skipping dist checks");
      return;
    }
    const files = walk(DIST);
    const fonts = files.filter((file) => FONT_PATTERN.test(file.path));
    expect(fonts).toEqual([]);
    const largest = [...files].sort((a, b) => b.size - a.size).slice(0, 10);
    const total = files.reduce((sum, file) => sum + file.size, 0);
    console.log(`web dist: ${files.length} files, ${(total / 1024 / 1024).toFixed(1)} MB total`);
    for (const file of largest) {
      console.log(`  ${(file.size / 1024 / 1024).toFixed(2)} MB  ${file.path}`);
    }
    expect(files.some((file) => file.path === "main.js")).toBe(true);
    expect(files.some((file) => file.path === "main.css")).toBe(true);
  });
});
