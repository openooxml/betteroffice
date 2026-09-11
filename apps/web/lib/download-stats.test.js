import { describe, expect, test } from "bun:test";
import { KV_KEYS, cachedDownloads, refreshDownloadStats } from "./download-stats.ts";
import { PYPI_PACKAGES } from "./pypi-downloads.ts";

const json = (body, status = 200) => ({ ok: status < 400, status, json: async () => body });

const NPM_LISTING = "https://registry.npmjs.org/-/org/betteroffice/package?format=cli";
const CRATES_SEARCH = "https://crates.io/api/v1/crates?page=1&per_page=100&q=betteroffice";
const today = "2026-09-07";

function history(downloads) {
  return { version_downloads: [{ date: today, downloads }], meta: { extra_downloads: [] } };
}

/** A registry table; every package answers 200 unless overridden. */
function registries(overrides = {}) {
  const table = {
    [NPM_LISTING]: () => json({ "@betteroffice/docx": "write", "@betteroffice/xlsx": "write", "other": "read" }),
    "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fdocx": () => json({ downloads: 100 }),
    "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fxlsx": () => json({ downloads: 20 }),
    [CRATES_SEARCH]: () =>
      json({
        crates: [
          { name: "betteroffice-opc", repository: "https://github.com/openooxml/betteroffice" },
          { name: "betteroffice-unrelated", repository: "https://github.com/example/x" },
        ],
      }),
    "https://crates.io/api/v1/crates/betteroffice-opc/downloads": () => json(history(7)),
  };
  for (const name of PYPI_PACKAGES) {
    table[`https://pypistats.org/api/packages/${encodeURIComponent(name)}/recent`] = () =>
      json({ data: { last_month: 3 } });
  }
  return { ...table, ...overrides };
}

function harness(table) {
  const calls = [];
  const sleeps = [];
  const store = { puts: [], async get() { return null; }, async put(key, value) { this.puts.push([key, JSON.parse(value)]); } };
  const fetchImpl = async (url) => {
    calls.push(url);
    const answer = table[url];
    if (!answer) throw new Error(`unexpected request ${url}`);
    return answer();
  };
  const options = { fetchImpl, sleep: async (ms) => { sleeps.push(ms); }, now: () => 1_700_000_000_000, intervalMs: 3000, attempts: 3 };
  return { calls, sleeps, store, options };
}

describe("scheduled download stats refresh", () => {
  test("requests one package at a time, three seconds apart, and writes every registry", async () => {
    const { calls, sleeps, store, options } = harness(registries());
    const result = await refreshDownloadStats(store, options);

    expect(result).toEqual({ npm: { downloads: 120 }, pypi: { downloads: 3 * PYPI_PACKAGES.length }, crates: { downloads: 7 } });
    expect(sleeps).toEqual(Array(calls.length - 1).fill(3000));
    expect(store.puts).toEqual([
      [KV_KEYS.npm, { downloads: 120, at: 1_700_000_000_000 }],
      [KV_KEYS.pypi, { downloads: 3 * PYPI_PACKAGES.length, at: 1_700_000_000_000 }],
      [KV_KEYS.crates, { downloads: 7, at: 1_700_000_000_000 }],
    ]);
  });

  test("retries a package that fails transiently, backing off before each retry", async () => {
    let attempts = 0;
    const flaky = () => {
      attempts += 1;
      return attempts < 3 ? json({}, 503) : json({ downloads: 20 });
    };
    const { calls, sleeps, store, options } = harness(
      registries({ "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fxlsx": flaky }),
    );
    const result = await refreshDownloadStats(store, options);

    expect(attempts).toBe(3);
    expect(result.npm).toEqual({ downloads: 120 });
    const expected = Array(calls.length - 1).fill(3000);
    expected.splice(2, 2, 6000, 12000);
    expect(sleeps).toEqual(expected);
  });

  test("waits for Retry-After when a registry rate-limits, capped at a minute", async () => {
    let attempts = 0;
    const limited = () => {
      attempts += 1;
      if (attempts === 1) return { ...json({}, 429), headers: new Headers({ "Retry-After": "7" }) };
      if (attempts === 2) return { ...json({}, 429), headers: new Headers({ "Retry-After": "3600" }) };
      return json({ downloads: 20 });
    };
    const { sleeps, store, options } = harness(
      registries({ "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fxlsx": limited }),
    );
    const result = await refreshDownloadStats(store, options);
    expect(result.npm).toEqual({ downloads: 120 });
    expect(sleeps.slice(2, 4)).toEqual([7000, 60000]);
  });

  test("does not write a registry whose package keeps failing, but still writes the others", async () => {
    const { store, options } = harness(
      registries({ "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fxlsx": () => json({}, 500) }),
    );
    const result = await refreshDownloadStats(store, options);

    expect(result.npm).toEqual({ error: expect.stringContaining("500") });
    expect(store.puts.map(([key]) => key)).toEqual([KV_KEYS.pypi, KV_KEYS.crates]);
  });

  test("a malformed count fails the registry instead of skewing the sum", async () => {
    const { store, options } = harness(
      registries({ "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fxlsx": () => json({ downloads: "many" }) }),
    );
    const result = await refreshDownloadStats(store, options);
    expect(result.npm).toEqual({ error: expect.stringContaining("invalid") });
    expect(store.puts.map(([key]) => key)).not.toContain(KV_KEYS.npm);
  });

  test("a package the stats API has never seen counts as zero", async () => {
    const { store, options } = harness(
      registries({
        "https://api.npmjs.org/downloads/point/last-month/%40betteroffice%2Fxlsx": () => json({}, 404),
        [`https://pypistats.org/api/packages/${encodeURIComponent(PYPI_PACKAGES[0])}/recent`]: () => json({}, 404),
      }),
    );
    const result = await refreshDownloadStats(store, options);
    expect(result.npm).toEqual({ downloads: 100 });
    expect(result.pypi).toEqual({ downloads: 3 * (PYPI_PACKAGES.length - 1) });
    expect(store.puts.map(([key]) => key)).toEqual([KV_KEYS.npm, KV_KEYS.pypi, KV_KEYS.crates]);
  });

  test("a failed listing fails only that registry", async () => {
    const { store, options } = harness(registries({ [CRATES_SEARCH]: () => json({}, 503) }));
    const result = await refreshDownloadStats(store, options);
    expect(result.crates).toEqual({ error: expect.stringContaining("503") });
    expect(store.puts.map(([key]) => key)).toEqual([KV_KEYS.npm, KV_KEYS.pypi]);
  });
});

describe("cached download counts", () => {
  test("serves what the refresh wrote and ignores garbage", async () => {
    const values = { [KV_KEYS.npm]: { downloads: 42, at: 1 }, [KV_KEYS.pypi]: { downloads: "x" } };
    const store = { async get(key) { return values[key] ?? null; }, async put() {} };
    expect(await cachedDownloads(store, "npm")).toEqual({ downloads: 42, at: 1 });
    expect(await cachedDownloads(store, "pypi")).toBeNull();
    expect(await cachedDownloads(store, "crates")).toBeNull();
  });
});
