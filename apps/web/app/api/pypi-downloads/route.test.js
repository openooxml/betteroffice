import { describe, expect, test } from "bun:test";
import { PYPI_PACKAGES, monthlyDownloads } from "../../../lib/pypi-downloads.ts";

const recent = (lastMonth, ok = true, status = 200) => ({
  ok,
  status,
  json: async () => ({ data: { last_month: lastMonth } }),
});

describe("PyPI downloads", () => {
  test("sums last-month downloads across packages", async () => {
    const byName = {
      "betteroffice-xlsx": recent(120),
      "betteroffice-docx": recent(30),
    };
    const result = await monthlyDownloads(Object.keys(byName), async (url) => {
      const name = decodeURIComponent(url.split("/packages/")[1].split("/")[0]);
      return byName[name];
    });
    expect(result).toEqual({ downloads: 150, resolved: 2 });
  });

  test("tolerates a package with no stats yet", async () => {
    // an unpublished distribution 404s until pypistats has data for it
    const byName = {
      "betteroffice-xlsx": recent(120),
      "betteroffice-docx": recent(undefined, false, 404),
    };
    const result = await monthlyDownloads(Object.keys(byName), async (url) => {
      const name = decodeURIComponent(url.split("/packages/")[1].split("/")[0]);
      return byName[name];
    });
    expect(result).toEqual({ downloads: 120, resolved: 1 });
  });

  test("reports nothing resolved when every package fails", async () => {
    const result = await monthlyDownloads(["betteroffice-xlsx"], async () =>
      recent(undefined, false, 404),
    );
    expect(result).toEqual({ downloads: 0, resolved: 0 });
  });

  test("rejects a malformed count rather than counting it", async () => {
    const result = await monthlyDownloads(["betteroffice-xlsx"], async () => ({
      ok: true,
      status: 200,
      json: async () => ({ data: { last_month: "many" } }),
    }));
    expect(result).toEqual({ downloads: 0, resolved: 0 });
  });

  test("every listed package is a betteroffice distribution", () => {
    expect(PYPI_PACKAGES.length).toBeGreaterThan(0);
    for (const name of PYPI_PACKAGES) {
      expect(name.startsWith("betteroffice-")).toBe(true);
    }
  });
});
