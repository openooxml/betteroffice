import { describe, expect, test } from "bun:test";
import { officialCrateNames, rollingDownloads } from "./route.ts";

describe("crates.io downloads", () => {
  test("keeps only BetterOffice crates from this repository", () => {
    expect(
      officialCrateNames([
        {
          name: "betteroffice-xlsx",
          repository: "https://github.com/openooxml/betteroffice.git",
        },
        {
          name: "betteroffice-opc",
          repository: "https://github.com/openooxml/betteroffice/",
        },
        {
          name: "betteroffice-unrelated",
          repository: "https://github.com/example/unrelated",
        },
        {
          name: "another-crate",
          repository: "https://github.com/openooxml/betteroffice",
        },
      ]),
    ).toEqual(["betteroffice-xlsx", "betteroffice-opc"]);
  });

  test("sums the latest 30 UTC dates across current and older versions", () => {
    expect(
      rollingDownloads(
        {
          version_downloads: [
            { date: "2026-06-18", downloads: 100 },
            { date: "2026-06-19", downloads: 2 },
            { date: "2026-07-18", downloads: 3 },
          ],
          meta: {
            extra_downloads: [{ date: "2026-07-01", downloads: 5 }],
          },
        },
        new Date("2026-07-18T18:00:00Z"),
      ),
    ).toBe(10);
  });

  test("rejects malformed download counts", () => {
    expect(() =>
      rollingDownloads({
        version_downloads: [{ date: "2026-07-18", downloads: -1 }],
        meta: { extra_downloads: [] },
      }),
    ).toThrow("Invalid crates.io download history");
  });
});
