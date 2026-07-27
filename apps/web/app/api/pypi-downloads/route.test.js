import { describe, expect, test } from "bun:test";
import {
  PYPI_PACKAGES,
  monthlyDownloads,
  monthlyDownloadsTotal,
} from "../../../lib/pypi-downloads.ts";

const recent = (lastMonth) => ({
  ok: true,
  status: 200,
  json: async () => ({ data: { last_month: lastMonth } }),
});
const status = (code) => ({ ok: false, status: code, json: async () => ({}) });

const byName = (table) => async (url) => {
  const name = decodeURIComponent(url.split("/packages/")[1].split("/")[0]);
  return table[name];
};

describe("PyPI downloads", () => {
  test("sums last-month downloads across packages", async () => {
    const result = await monthlyDownloads(
      ["betteroffice-xlsx", "betteroffice-docx"],
      byName({
        "betteroffice-xlsx": recent(120),
        "betteroffice-docx": recent(30),
      }),
    );
    expect(result).toEqual({ downloads: 150, resolved: 2, missing: 0 });
  });

  test("counts a package pypistats has never seen as missing, not failed", async () => {
    const result = await monthlyDownloads(
      ["betteroffice-xlsx", "betteroffice-docx"],
      byName({
        "betteroffice-xlsx": recent(120),
        "betteroffice-docx": status(404),
      }),
    );
    expect(result).toEqual({ downloads: 120, resolved: 1, missing: 1 });
  });

  test("a failing API is not counted as missing", async () => {
    const result = await monthlyDownloads(["betteroffice-xlsx"], async () =>
      status(429),
    );
    expect(result).toEqual({ downloads: 0, resolved: 0, missing: 0 });
  });

  test("rejects a malformed count rather than counting it", async () => {
    const result = await monthlyDownloads(["betteroffice-xlsx"], async () => ({
      ok: true,
      status: 200,
      json: async () => ({ data: { last_month: "many" } }),
    }));
    expect(result).toEqual({ downloads: 0, resolved: 0, missing: 0 });
  });
});

describe("PyPI downloads total", () => {
  test("an unpublished package reads as zero, not unavailable", async () => {
    expect(await monthlyDownloadsTotal(async () => status(404))).toBe(0);
  });

  test("a rate-limited or failing API reads as unavailable", async () => {
    expect(await monthlyDownloadsTotal(async () => status(429))).toBeNull();
    expect(await monthlyDownloadsTotal(async () => status(503))).toBeNull();
    expect(
      await monthlyDownloadsTotal(async () => {
        throw new Error("network down");
      }),
    ).toBeNull();
  });

  test("reports the count once a package has stats", async () => {
    expect(await monthlyDownloadsTotal(async () => recent(42))).toBe(42);
  });

  test("every listed package is a betteroffice distribution", () => {
    expect(PYPI_PACKAGES.length).toBeGreaterThan(0);
    for (const name of PYPI_PACKAGES) {
      expect(name.startsWith("betteroffice-")).toBe(true);
    }
  });
});
