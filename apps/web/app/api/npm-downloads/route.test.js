import { describe, expect, test } from "bun:test";
import { monthlyDownloads, officialPackageNames } from "../../../lib/downloads.ts";

const jsonResponse = (body, ok = true, status = 200) => ({
  ok,
  status,
  json: async () => body,
});

describe("npm downloads", () => {
  test("keeps only @betteroffice packages", () => {
    expect(
      officialPackageNames({
        "@betteroffice/docx": "write",
        "@betteroffice/xlsx": "write",
        "other-package": "read",
      }),
    ).toEqual(["@betteroffice/docx", "@betteroffice/xlsx"]);
  });

  test("tolerates per-package failures and sums the rest", async () => {
    const byName = {
      "@betteroffice/xlsx": jsonResponse({ downloads: 120 }),
      // freshly published packages 404 on the stats API for up to a day
      "@betteroffice/docx": jsonResponse({}, false, 404),
      "@betteroffice/docx-react": jsonResponse({ downloads: "bad" }),
      "@betteroffice/xlsx-react": jsonResponse({ downloads: 30 }),
    };
    const fetchImpl = async (url) => {
      const name = decodeURIComponent(String(url).split("/last-month/")[1]);
      return byName[name];
    };

    expect(
      await monthlyDownloads(Object.keys(byName), fetchImpl),
    ).toEqual({ downloads: 150, resolved: 2 });
  });

  test("reports zero resolved when every fetch fails", async () => {
    const fetchImpl = async () => jsonResponse({}, false, 500);
    expect(
      await monthlyDownloads(["@betteroffice/docx"], fetchImpl),
    ).toEqual({ downloads: 0, resolved: 0 });
  });
});
