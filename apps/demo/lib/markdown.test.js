import { describe, expect, test } from "bun:test";
import { formats } from "./formats.ts";
import { formatMarkdown, indexMarkdown } from "./markdown.ts";

describe("demo markdown", () => {
  test("index links every format", () => {
    const markdown = indexMarkdown();
    for (const format of formats) {
      expect(markdown).toContain(`/${format.id}`);
      expect(markdown).toContain(format.tagline);
    }
  });

  test("each format has a page", () => {
    for (const format of formats) {
      const markdown = formatMarkdown(format.id);
      expect(markdown).toContain(`# BetterOffice ${format.id.toUpperCase()} demo`);
      expect(markdown).toContain(`@betteroffice/${format.id}`);
    }
  });

  test("an unknown format has none", () => {
    expect(formatMarkdown("odt")).toBeNull();
    expect(formatMarkdown("")).toBeNull();
  });

  test("carries no HTML tags", () => {
    expect(indexMarkdown()).not.toMatch(/<[a-z][^>]*>/i);
  });
});
