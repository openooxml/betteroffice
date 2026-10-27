import { BENCHMARKS_SITE, DEMO, DOCS, EVALS_SITE, SITE } from "../../../shared/sites";

const escape = (value: string) => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

export const SITEMAP_HEADERS = {
  "Content-Type": "application/xml; charset=utf-8",
  "Cache-Control": "public, max-age=3600",
};

export const CHILD_SITEMAPS = [`${SITE}/sitemap-pages.xml`, `${DOCS}/sitemap.xml`, `${DEMO}/sitemap.xml`, `${EVALS_SITE}/sitemap.xml`, `${BENCHMARKS_SITE}/sitemap.xml`];

export function sitemapIndexXml(children = CHILD_SITEMAPS) {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${children.map((loc) => `  <sitemap><loc>${escape(loc)}</loc></sitemap>`).join("\n")}\n</sitemapindex>\n`;
}

export function urlsetXml(entries: { url: string; lastModified?: string | Date; changeFrequency?: string; priority?: number }[]) {
  const date = (value: string | Date) => (value instanceof Date ? value.toISOString() : value);
  return `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${entries
    .map(
      (entry) =>
        `  <url><loc>${escape(entry.url)}</loc>${entry.lastModified ? `<lastmod>${escape(date(entry.lastModified))}</lastmod>` : ""}${entry.changeFrequency ? `<changefreq>${entry.changeFrequency}</changefreq>` : ""}${entry.priority === undefined ? "" : `<priority>${entry.priority}</priority>`}</url>`,
    )
    .join("\n")}\n</urlset>\n`;
}
