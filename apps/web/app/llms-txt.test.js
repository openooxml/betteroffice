import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import {
  PYPI_DISTRIBUTIONS,
  publishedCrates,
  publishedPackages,
} from "../../../scripts/published-packages.mjs";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { agentEvalCards, agentEvalTasks, parseAgentEvals } from "../lib/agent-evals.ts";
import { BENCHMARKS, BENCHMARKS_SITE, DEMO, DOCS, EDITORS, PACKAGES, RELEASES, REPO, SITE } from "./content.ts";
import { USED_BY, usedByLine } from "./used-by.ts";

const ROOT = join(import.meta.dir, "..", "..", "..");
const LLMS = readFileSync(join(ROOT, "apps/web/content/llms.txt"), "utf8");

const README = readFileSync(join(ROOT, "README.md"), "utf8");
process.env.BETTEROFFICE_BENCHMARKS = README;
process.env.BETTEROFFICE_AGENT_EVALS = README;
const EVALS = parseAgentEvals(README);
const { default: Home } = await import("./page.tsx");
const PAGE = renderToStaticMarkup(createElement(Home));

// what llms.txt carries, at index level, for each section of the landing page
const COVERAGE = {
  hero: ["office suite", "written in Rust", "agent", `(${DEMO})`, `(${DOCS})`, `(${REPO})`],
  "used-by": [usedByLine()],
  suite: [...EDITORS.map((editor) => editor.format.toUpperCase()), "published on npm", "source preview", "edited in their own format"],
  foundation: ["tracked changes", "undoes as one step", "Untouched file parts", "self-hostable", `(${DOCS}/docs)`],
  benchmarks: [
    `(${BENCHMARKS})`, `(${BENCHMARKS_SITE})`, `(${SITE}/#benchmarks)`, "LibreOffice", "render and recalculation time", "scripts/office-quality",
    `(${REPO}#evals)`, `${agentEvalTasks(EVALS)} editing tasks`,
    ...agentEvalCards(EVALS).map((card) => `${card.format.toUpperCase()} ${card.bars[0].text} vs ${card.bars[1].text}`),
  ],
  collaboration: ["People and agents", "offline edits", `(${DOCS}/docs/collaboration)`],
  packages: ["never leave the page", ...PACKAGES.map((pkg) => `[${pkg.name}](`), `(${DOCS}/docs/packages)`, `(${RELEASES})`],
  coda: [`(${DEMO})`, `(${DOCS})`],
};

describe("llms.txt", () => {
  test("covers every section of the landing page", () => {
    const sections = [...PAGE.matchAll(/<section[^>]*aria-labelledby="([^"]+)"/g)].map((m) => m[1]);
    expect(sections).toContain("benchmarks");
    expect(sections.filter((id) => !(id in COVERAGE))).toEqual([]);
    for (const id of sections) {
      expect({ [id]: COVERAGE[id].filter((text) => !LLMS.includes(text)) }).toEqual({ [id]: [] });
    }
  });

  test("names who uses it exactly as the band does", () => {
    expect(LLMS.split("\n")).toContain(usedByLine());
    for (const entry of USED_BY.filter((e) => !e.enabled)) expect(LLMS).not.toContain(entry.name);
  });

  test("names no npm package that does not exist", () => {
    const real = new Set(publishedPackages());
    const claimed = [...LLMS.matchAll(/@betteroffice\/[a-z0-9-]+/g)].map((m) => m[0]);
    const phantom = [...new Set(claimed)].filter((name) => !real.has(name));
    expect(phantom).toEqual([]);
  });

  test("links every published npm package", () => {
    const missing = publishedPackages().filter(
      (name) => !LLMS.includes(`[${name}](https://www.npmjs.com/package/${name})`),
    );
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

  test("follows the llms.txt index layout", () => {
    const lines = LLMS.split("\n");
    expect(lines[0]).toBe("# BetterOffice");
    expect(lines[2].startsWith("> ")).toBe(true);
    for (const section of ["## Docs", "## npm packages", "## Rust crates", "## Python packages"]) {
      expect(lines).toContain(section);
    }
  });

  test("keeps every entry to a short description", () => {
    const entries = LLMS.split("\n").filter((line) => line.startsWith("- ["));
    const long = entries.filter((line) => {
      const desc = line.slice(line.lastIndexOf("): ") + 3);
      return desc.split(/\s+/).length > 45 || desc.split(/\.\s/).length > 2;
    });
    expect(long).toEqual([]);
  });

  test("tells no one to install a distribution that is not on PyPI", () => {
    const real = new Set(PYPI_DISTRIBUTIONS);
    const claimed = [...LLMS.matchAll(/pip install (betteroffice-[a-z0-9-]+)/g)].map(
      (m) => m[1],
    );
    expect(claimed.filter((name) => !real.has(name))).toEqual([]);
  });
});
