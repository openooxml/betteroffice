import { describe, expect, test } from "bun:test";
import { prefersMarkdown } from "../../../shared/accept.ts";
import { homepageMarkdown } from "./markdown.ts";
import { EDITORS, HERO, PACKAGES } from "./content.ts";

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

  test("carries no HTML tags", () => {
    expect(markdown).not.toMatch(/<[a-z][^>]*>/i);
  });
});
