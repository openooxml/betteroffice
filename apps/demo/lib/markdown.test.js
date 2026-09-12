import { describe, expect, test } from "bun:test";
import { formats, liveFormats } from "./formats.ts";
import { formatMarkdown, indexMarkdown } from "./markdown.ts";

describe("demo markdown", () => {
  test("index links every live format", () => {
    const markdown = indexMarkdown();
    for (const format of liveFormats) {
      expect(markdown).toContain(`/${format.id}`);
      expect(markdown).toContain(format.tagline);
    }
    for (const format of formats.filter((format) => format.status !== "live")) {
      expect(markdown).not.toContain(`/${format.id}`);
    }
  });

  test("each live format has a page", () => {
    for (const format of liveFormats) {
      const markdown = formatMarkdown(format.id);
      expect(markdown).toContain(`# BetterOffice ${format.id.toUpperCase()} demo`);
      expect(markdown).toContain(`@betteroffice/${format.id}`);
    }
  });

  test("a non-live format has no page", () => {
    expect(formatMarkdown("vsdx")).toBeNull();
  });

  test("an unknown format has none", () => {
    expect(formatMarkdown("odt")).toBeNull();
    expect(formatMarkdown("")).toBeNull();
  });

  test("carries no HTML tags", () => {
    expect(indexMarkdown()).not.toMatch(/<[a-z][^>]*>/i);
  });
});
