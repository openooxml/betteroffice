import { describe, expect, test } from "bun:test";
import { formats, liveFormats, listedLiveFormats } from "./formats.ts";
import { formatMarkdown, indexMarkdown } from "./markdown.ts";

describe("demo markdown", () => {
  test("index links every listed live format", () => {
    const markdown = indexMarkdown();
    for (const format of listedLiveFormats) {
      expect(markdown).toContain(`/${format.id}`);
      expect(markdown).toContain(format.tagline);
    }
    for (const format of formats.filter((format) => !listedLiveFormats.includes(format))) {
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

  test("describes the VSDX route as an editor", () => {
    const markdown = formatMarkdown("vsdx");
    expect(markdown).toContain("interactive editor");
    expect(markdown).toContain("React editor");
  });

  test("a non-live format has no page", () => {
    expect(formatMarkdown("odt")).toBeNull();
  });

  test("an unknown format has none", () => {
    expect(formatMarkdown("drawio")).toBeNull();
    expect(formatMarkdown("")).toBeNull();
  });

  test("carries no HTML tags", () => {
    expect(indexMarkdown()).not.toMatch(/<[a-z][^>]*>/i);
  });
});
