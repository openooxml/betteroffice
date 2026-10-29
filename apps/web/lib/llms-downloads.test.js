import { describe, expect, test } from "bun:test";
import { downloadsLine, liveDownloadsLine, withDownloads } from "./llms-downloads.ts";

const AT = Date.UTC(2026, 9, 27, 3, 0, 0);
const BASE = "# BetterOffice\n\n> Summary.\n\nIntro paragraph.\n\n## Features\n\nText.\n\n## Links\n";

describe("llms.txt download line", () => {
  test("lists each registry and the date of the newest refresh", () => {
    const line = downloadsLine({
      npm: { downloads: 484591, at: AT },
      crates: { downloads: 203294, at: AT - 86_400_000 },
      pypi: { downloads: 2063, at: AT },
    });
    expect(line).toBe("Downloads in the last 30 days: npm 484,591, crates.io 203,294, PyPI 2,063 (as of 2026-10-27).");
  });

  test("names only the registries that have a count", () => {
    expect(downloadsLine({ npm: { downloads: 5, at: AT }, crates: null })).toBe("Downloads in the last 30 days: npm 5 (as of 2026-10-27).");
  });

  test("is empty without any count, and then llms.txt is unchanged", () => {
    expect(downloadsLine({})).toBe("");
    expect(downloadsLine({ npm: null, crates: null, pypi: null })).toBe("");
    expect(withDownloads(BASE, "")).toBe(BASE);
  });

  test("sits after the introduction and before the first section, once", () => {
    const line = downloadsLine({ npm: { downloads: 1, at: AT } });
    const text = withDownloads(BASE, line);
    expect(text.split("\n").filter((l) => l.startsWith("Downloads in the last 30 days"))).toHaveLength(1);
    expect(text.indexOf(line)).toBeGreaterThan(text.indexOf("Intro paragraph."));
    expect(text.indexOf(line)).toBeLessThan(text.indexOf("## Features"));
    expect(text).toContain(`${line}\n\n## Features`);
  });

  test("a store with nothing in it, or none, yields no line", async () => {
    expect(await liveDownloadsLine(async () => null)).toBe("");
  });

  test("reads the three cached counts", async () => {
    const stored = { npm: { downloads: 10, at: AT }, crates: { downloads: 20, at: AT }, pypi: { downloads: 30, at: AT } };
    expect(await liveDownloadsLine(async (registry) => stored[registry])).toBe("Downloads in the last 30 days: npm 10, crates.io 20, PyPI 30 (as of 2026-10-27).");
  });
});
