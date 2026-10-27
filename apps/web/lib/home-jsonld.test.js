import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { downloadStatistics, liveCounts, softwareJsonLd } from "./home-jsonld.ts";

const README = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");
process.env.BETTEROFFICE_BENCHMARKS = README;
process.env.BETTEROFFICE_AGENT_EVALS = README;
const { Home } = await import("../app/home.tsx");

const AT = Date.UTC(2026, 9, 27, 3, 0, 0);
const COUNTS = { npm: { downloads: 484591, at: AT }, crates: { downloads: 203294, at: AT - 3_600_000 }, pypi: { downloads: 2063, at: AT } };

const blockIn = (html) => {
  const blocks = [...html.matchAll(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/g)].map((m) => JSON.parse(m[1].replaceAll("&quot;", '"')));
  return blocks.find((block) => block["@type"] === "SoftwareApplication");
};

describe("homepage SoftwareApplication JSON-LD", () => {
  test("is parseable, typed and carries the fields Google reads for software", () => {
    const data = blockIn(renderToStaticMarkup(createElement(Home, { downloads: COUNTS })));
    expect(data["@context"]).toBe("https://schema.org");
    expect(data["@type"]).toBe("SoftwareApplication");
    expect(data.name).toBe("BetterOffice");
    expect(data.applicationCategory).toBe("BusinessApplication");
    expect(data.operatingSystem).toBeTruthy();
    expect(data.offers).toMatchObject({ "@type": "Offer", price: "0", priceCurrency: "USD" });
  });

  test("states no rating or review, so it is structured data and not a rich-result candidate", () => {
    const data = softwareJsonLd(COUNTS);
    expect(data.aggregateRating).toBeUndefined();
    expect(data.review).toBeUndefined();
  });

  test("download counters equal the stored counts, in total and per registry", () => {
    const data = blockIn(renderToStaticMarkup(createElement(Home, { downloads: COUNTS })));
    const [total, ...each] = data.interactionStatistic;
    for (const counter of data.interactionStatistic) {
      expect(counter["@type"]).toBe("InteractionCounter");
      expect(counter.interactionType).toEqual({ "@type": "DownloadAction" });
      expect(counter.name).toBeTruthy();
      expect(counter.description).toBeTruthy();
    }
    expect(total.userInteractionCount).toBe(484591 + 203294 + 2063);
    expect(each.map((counter) => [counter.interactionService.name, counter.userInteractionCount])).toEqual([
      ["npm", 484591],
      ["crates.io", 203294],
      ["PyPI", 2063],
    ]);
  });

  test("dateModified is the newest refresh as an ISO timestamp", () => {
    const data = softwareJsonLd(COUNTS);
    expect(data.dateModified).toBe("2026-10-27T03:00:00.000Z");
    expect(new Date(data.dateModified).toISOString()).toBe(data.dateModified);
  });

  test("a registry without a count is left out, and an empty store drops only the counters", () => {
    const partial = downloadStatistics({ npm: { downloads: 7, at: AT }, crates: null });
    expect(partial.interactionStatistic.map((counter) => counter.userInteractionCount)).toEqual([7, 7]);
    for (const empty of [{}, { npm: null, crates: null, pypi: null }]) {
      const data = blockIn(renderToStaticMarkup(createElement(Home, { downloads: empty })));
      expect(data["@type"]).toBe("SoftwareApplication");
      expect(data.name).toBe("BetterOffice");
      expect(data.interactionStatistic).toBeUndefined();
      expect(data.dateModified).toBeUndefined();
    }
  });

  test("reads the three cached registries", async () => {
    expect(await liveCounts(async (registry) => COUNTS[registry])).toEqual(COUNTS);
    expect(await liveCounts(async () => null)).toEqual({ npm: null, crates: null, pypi: null });
  });
});
