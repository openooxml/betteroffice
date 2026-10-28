import { describe, expect, test } from "bun:test";
import { CHILD_SITEMAPS, sitemapIndexXml, urlsetXml } from "./sitemap-xml.ts";

describe("sitemap index", () => {
  test("references the site, docs, demo, evals and benchmarks sitemaps", () => {
    const xml = sitemapIndexXml();
    expect(xml).toContain('<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">');
    for (const loc of [
      "https://betteroffice.dev/sitemap-pages.xml",
      "https://docs.betteroffice.dev/sitemap.xml",
      "https://demo.betteroffice.dev/sitemap.xml",
      "https://evals.betteroffice.dev/sitemap.xml",
      "https://benchmarks.betteroffice.dev/sitemap.xml",
    ]) expect(xml).toContain(`<loc>${loc}</loc>`);
    expect(xml.match(/<sitemap>/g)).toHaveLength(CHILD_SITEMAPS.length);
  });

  test("the pages sitemap keeps the existing URLs and escapes them", () => {
    const xml = urlsetXml([{ url: "https://betteroffice.dev", changeFrequency: "weekly", priority: 1 }, { url: "https://betteroffice.dev/a?b=1&c=2", lastModified: "2026-10-10" }]);
    expect(xml).toContain("<loc>https://betteroffice.dev</loc><changefreq>weekly</changefreq><priority>1</priority>");
    expect(xml).toContain("<loc>https://betteroffice.dev/a?b=1&amp;c=2</loc><lastmod>2026-10-10</lastmod>");
  });

  test("the routes serve the index and the pages list as XML", async () => {
    const index = await (await import("../app/sitemap.xml/route.ts")).GET();
    expect(index.headers.get("content-type")).toContain("application/xml");
    expect(await index.text()).toContain("<sitemapindex");
    const pages = await (await import("../app/sitemap-pages.xml/route.ts")).GET();
    expect(await pages.text()).toContain("<loc>https://betteroffice.dev</loc>");
  });
});
