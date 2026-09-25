import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import {
  PYPI_DISTRIBUTIONS,
  publishedCrates,
  publishedPackages,
} from "../../../scripts/published-packages.mjs";

const ROOT = join(import.meta.dir, "..", "..", "..");
const LLMS = readFileSync(join(ROOT, "apps/web/public/llms.txt"), "utf8");

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
    // A PyPI distribution shares the crate naming shape without being a crate.
    const pypi = new Set(PYPI_DISTRIBUTIONS);
    const claimed = [...LLMS.matchAll(/\bbetteroffice(?:-[a-z0-9]+)+\b/g)].map((m) => m[0]);
    const phantom = [...new Set(claimed)].filter(
      (name) => !real.has(name) && !pypi.has(name),
    );
    expect(phantom).toEqual([]);
  });

  test("names every published crate", () => {
    const missing = publishedCrates().filter((name) => !LLMS.includes(name));
    expect(missing).toEqual([]);
  });

  test("links every published PyPI distribution", () => {
    // A bare name would already match the same-named crate.
    const missing = PYPI_DISTRIBUTIONS.filter(
      (name) => !LLMS.includes(`https://pypi.org/project/${name}`),
    );
    expect(missing).toEqual([]);
  });

  test("states what XLSX structured export covers and that it never recalculates", () => {
    expect(LLMS).toContain(
      "XLSX exports bounded sparse worksheet content and Markdown with positional anchors, formulas, stored values, formatted text, explicit hidden-content options, and omission diagnostics",
    );
    expect(LLMS).toContain("Export does not recalculate formulas.");
    for (const api of [
      "exportStructured",
      "exportMarkdown",
      "exportXlsxStructured",
      "exportXlsxMarkdown",
      "renderXlsxMarkdown",
      "export_structured",
    ]) {
      expect(LLMS).toContain(api);
    }
  });

  test("tells no one to install a distribution that is not on PyPI", () => {
    const real = new Set(PYPI_DISTRIBUTIONS);
    const claimed = [...LLMS.matchAll(/pip install (betteroffice-[a-z0-9-]+)/g)].map(
      (m) => m[1],
    );
    expect(claimed.filter((name) => !real.has(name))).toEqual([]);
  });
});
