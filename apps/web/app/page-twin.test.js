import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { parseFidelity } from "../lib/fidelity.ts";
import { parseAgentEvals } from "../lib/agent-evals.ts";
import { homepageMarkdown } from "./markdown.ts";
import { USED_BY, USED_BY_LABEL, USED_BY_MORE } from "./used-by.ts";

const README = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");
process.env.BETTEROFFICE_BENCHMARKS = README;
process.env.BETTEROFFICE_AGENT_EVALS = README;
const { default: Home } = await import("./page.tsx");

const html = renderToStaticMarkup(createElement(Home));
const markdown = homepageMarkdown(parseFidelity(README), parseAgentEvals(README));

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', "#x27": "'", "#39": "'" };
const plain = (fragment) =>
  fragment
    .replace(/<[^>]+>/g, "")
    .replace(/&(amp|lt|gt|quot|#x27|#39);/g, (_, name) => ENTITIES[name])
    .replace(/\s+/g, " ")
    .trim();
const elements = (tags) =>
  [...html.matchAll(new RegExp(`<(${tags})\\b[^>]*>([\\s\\S]*?)</\\1>`, "g"))].map((m) => plain(m[2]));
const read = markdown
  .replace(/\[([^\]]*)\]\([^)]*\)/g, "$1")
  .replace(/\*\*|`/g, "")
  .replace(/\s+/g, " ");

describe("markdown twin of the page", () => {
  test("renders the benchmark section it is checked against", () => {
    expect(elements("h2")).toContain("Benchmarks");
    expect(elements("h3")).toContain("Agent edits");
  });

  test("carries every heading on the page", () => {
    for (const heading of elements("h1|h2|h3|h4")) expect(read).toContain(heading);
  });

  test("carries every word the benchmark cards print, in the cards' order", () => {
    const section = html.slice(html.indexOf('aria-labelledby="benchmarks"'));
    const cards = section
      .slice(section.indexOf("<ul"), section.indexOf("</section>"))
      .replace(/<(\w+)[^>]*(?:aria-hidden="true"|class="sr-only")[^>]*>[\s\S]*?<\/\1>/g, "");
    const texts = cards
      .split(/<[^>]+>/)
      .map(plain)
      .filter(Boolean);
    expect(texts).toContain("Render time");
    expect(texts).toContain("Recalculation time");
    for (const text of texts) expect(read).toContain(text);
    let from = 0;
    for (const label of elements("h4")) {
      from = read.indexOf(label, from);
      expect(from).toBeGreaterThanOrEqual(0);
      from += label.length;
    }
  });

  test("carries every sentence the page prints as a paragraph", () => {
    const sentences = elements("p").filter((text) => /[.!?]$/.test(text));
    expect(sentences.length).toBeGreaterThan(10);
    for (const sentence of sentences) expect(read).toContain(sentence);
  });

  test("prints the used-by names as text, not only as logos", () => {
    const start = html.indexOf('<section aria-labelledby="used-by"');
    const band = html.slice(start, html.indexOf("</section>", start));
    const names = USED_BY.filter((entry) => entry.enabled).map((entry) => entry.name);
    expect(band.split(/<[^>]+>/).map(plain).filter(Boolean)).toEqual([USED_BY_LABEL, ...names, USED_BY_MORE]);
    expect(band).not.toMatch(/aria-label=/);
  });

  test("keeps the page's section order", () => {
    const at = elements("h2").map((heading) =>
      markdown.search(new RegExp(`^(## )?${heading.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "m")),
    );
    expect(at.every((index) => index >= 0)).toBe(true);
    expect(at).toEqual([...at].sort((a, b) => a - b));
  });
});
