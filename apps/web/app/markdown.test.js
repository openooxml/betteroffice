import { describe, expect, test } from "bun:test";
import { prefersMarkdown } from "../../../shared/accept.ts";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fidelityCharts, parseFidelity } from "../lib/fidelity.ts";
import { agentEvalCards, agentEvalRuns, parseAgentEvals } from "../lib/agent-evals.ts";
import { homepageMarkdown } from "./markdown.ts";
import { AGENT_EVALS, BENCHMARKS_SECTION, BENCHMARKS_SITE, DEMO, ECOSYSTEMS, EDITORS, editorOrder, HERO, PACKAGES } from "./content.ts";
import { USED_BY, USED_BY_LABEL, USED_BY_MORE, usedByLine } from "./used-by.ts";

describe("accept negotiation", () => {
  test("serves HTML to browsers", () => {
    expect(
      prefersMarkdown(
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,*/*;q=0.8",
      ),
    ).toBe(false);
    expect(prefersMarkdown(null)).toBe(false);
    expect(prefersMarkdown("*/*")).toBe(false);
  });

  test("serves markdown when an agent asks for it", () => {
    expect(prefersMarkdown("text/markdown")).toBe(true);
    expect(prefersMarkdown("text/markdown, text/html;q=0.5")).toBe(true);
    expect(prefersMarkdown("TEXT/MARKDOWN")).toBe(true);
  });

  test("honours wildcard ranges", () => {
    // text/* gives HTML an effective 0.9, above markdown's explicit 0.5
    expect(prefersMarkdown("text/*;q=0.9, text/markdown;q=0.5")).toBe(false);
    expect(prefersMarkdown("text/*;q=0.5, text/markdown;q=0.9")).toBe(true);
    expect(prefersMarkdown("*/*;q=0.8, text/markdown")).toBe(true);
    expect(prefersMarkdown("*/*, text/markdown;q=0.5")).toBe(false);
    expect(prefersMarkdown("text/*")).toBe(false);
  });

  test("respects quality values in both directions", () => {
    expect(prefersMarkdown("text/markdown;q=0.9,text/html;q=1.0")).toBe(false);
    expect(prefersMarkdown("text/markdown;q=1.0,text/html;q=0.9")).toBe(true);
    expect(prefersMarkdown("text/markdown;q=0")).toBe(false);
  });
});

describe("homepage markdown", () => {
  const markdown = homepageMarkdown();

  test("leads with the same title and tagline as the page", () => {
    expect(markdown.startsWith(`# ${HERO.title}\n`)).toBe(true);
    expect(markdown).toContain(HERO.tagline);
  });

  test("covers every editor and package the page renders", () => {
    for (const editor of EDITORS) {
      expect(markdown).toContain(editor.name);
      expect(markdown).toContain(editor.desc);
    }
    for (const pkg of PACKAGES) {
      expect(markdown).toContain(pkg.name);
      expect(markdown).toContain(pkg.desc);
    }
  });

  test("names who uses it, from the band's data", () => {
    expect(markdown).toContain(`\n${usedByLine()}\n`);
    expect(usedByLine().startsWith(`${USED_BY_LABEL}: `)).toBe(true);
    expect(usedByLine().endsWith(` ${USED_BY_MORE}`)).toBe(true);
    for (const entry of USED_BY) {
      expect(markdown.includes(`[${entry.name}](${entry.href})`)).toBe(entry.enabled);
    }
  });

  test("links a demo only for editors that have one", () => {
    for (const editor of EDITORS) {
      expect(markdown.includes(`(${DEMO}/${editor.format})`)).toBe(editor.demo);
    }
    expect(EDITORS.find((editor) => editor.format === "vsdx").demo).toBe(false);
  });

  test("describes every package in one short sentence", () => {
    const long = PACKAGES.filter(
      (pkg) => pkg.desc.split(/\s+/).length > 25 || pkg.desc.split(/\.\s/).length > 1,
    );
    expect(long).toEqual([]);
  });

  test("carries an install line and a guide for every ecosystem", () => {
    for (const eco of ECOSYSTEMS) {
      expect(markdown).toContain(eco.install);
      expect(markdown).toContain(eco.docs);
    }
  });

  test("carries no HTML tags", () => {
    expect(markdown).not.toMatch(/<[a-z][^>]*>/i);
  });
});

describe("homepage markdown benchmarks", () => {
  const formats = parseFidelity(readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8"));
  const markdown = homepageMarkdown(formats);

  const card = (f) => markdown.split("\n### ").find((section) => section.includes(`(.${f.format})\n`));

  test("carries every published row of the README's benchmark tables", () => {
    expect(markdown).toContain("## Benchmarks");
    for (const f of formats) {
      expect(card(f)).toContain(`All results:\n\n| | BetterOffice ([${f.commit}](${f.commitUrl})) | LibreOffice (${f.libreoffice}) |`);
      for (const row of f.rows) expect(card(f)).toContain(`| ${row.label} | ${row.ours} | ${row.theirs} |`);
    }
  });

  test("lists each card's quality rows, then its time rows, before the full table", () => {
    for (const f of formats) {
      const lines = card(f).split("\n");
      const at = (chart) => lines.findIndex((line) => line.startsWith(`- **${chart.label}**, ${chart.hint} (`));
      const drawn = fidelityCharts(f);
      const quality = drawn.filter((chart) => chart.group === "quality").map(at);
      const time = drawn.filter((chart) => chart.group === "time").map(at);
      expect(time.length).toBeGreaterThan(0);
      expect([...quality, ...time].every((index) => index > 0)).toBe(true);
      expect(Math.min(...time)).toBeGreaterThan(Math.max(...quality) + 1);
      expect(Math.max(...time)).toBeLessThan(lines.indexOf("All results:"));
      for (const chart of drawn) {
        expect(lines[at(chart)]).toEndWith(`: BetterOffice ${chart.bars[0].text}, LibreOffice ${chart.bars[1].text}`);
      }
    }
  });

  test("names the measured builds and links the per-document results", () => {
    const [f] = formats;
    expect(markdown).toContain(`\nBetterOffice [${f.commit}](${f.commitUrl}), LibreOffice ${f.libreoffice}\n`);
    expect(markdown).toContain(`[${BENCHMARKS_SECTION.more}](${BENCHMARKS_SITE})`);
  });

  test("leaves the section out when there are no results", () => {
    expect(homepageMarkdown()).not.toContain("## Benchmarks");
  });

  test("stays free of HTML with the tables in", () => {
    expect(markdown).not.toMatch(/<[a-z][^>]*>/i);
  });
});

describe("homepage markdown agent evals", () => {
  const readme = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");
  const formats = parseFidelity(readme);
  const evals = parseAgentEvals(readme);
  const markdown = homepageMarkdown(formats, evals);

  test("carries the heading, graded run total and method link", () => {
    expect(markdown).toContain("### Agent edits\n");
    expect(markdown).toContain(AGENT_EVALS.lead(agentEvalRuns(evals)));
    expect(markdown).toContain(`${agentEvalRuns(evals)} graded runs`);
    expect(markdown).toContain(`[${AGENT_EVALS.more}](${AGENT_EVALS.method})`);
  });

  test("prints every drawn card in suite order with the published labels and values", () => {
    const cards = agentEvalCards(evals).sort((a, b) => editorOrder(a.format) - editorOrder(b.format));
    const lines = cards.map((card) => {
      const name = EDITORS.find((editor) => editor.format === card.format).name;
      return `- **${name}** (.${card.format}): ${evals.ours} ${card.bars[0].text}, ${evals.theirs} ${card.bars[1].text}`;
    });
    expect(markdown).toContain(lines.join("\n"));
    expect(homepageMarkdown(formats, { ...evals, rows: [{ format: "docx", ours: "n/a", theirs: "6/10" }] })).not.toContain("### Agent edits");
  });

  test("sits after the fidelity results and before collaboration", () => {
    const start = markdown.indexOf("### Agent edits");
    expect(start).toBeGreaterThan(markdown.indexOf(`[${BENCHMARKS_SECTION.more}](${BENCHMARKS_SITE})`));
    expect(start).toBeLessThan(markdown.indexOf("## People and agents, one document"));
  });

  test("leaves the block out without evals or fidelity", () => {
    expect(homepageMarkdown(formats)).not.toContain("### Agent edits");
    expect(homepageMarkdown(formats, null)).not.toContain("### Agent edits");
    expect(homepageMarkdown([], evals)).not.toContain("### Agent edits");
  });

  test("stays free of HTML with the evals in", () => {
    expect(markdown).not.toMatch(/<[a-z][^>]*>/i);
  });
});
