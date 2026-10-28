import { mkdir } from 'node:fs/promises';

export const ORIGIN = 'https://benchmarks.betteroffice.dev';
export const PAGES = ['/', '/compare'];

export const robots = () => `User-agent: *\nAllow: /\n\nSitemap: ${ORIGIN}/sitemap.xml\n`;

export const sitemap = (date: string) =>
  `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${PAGES.map(
    (path) => `  <url><loc>${ORIGIN}${path}</loc><lastmod>${date}</lastmod></url>`,
  ).join('\n')}\n</urlset>\n`;

if (import.meta.main) {
  const dist = `${import.meta.dir}/../dist`;
  await mkdir(dist, { recursive: true });
  await Bun.write(`${dist}/robots.txt`, robots());
  await Bun.write(`${dist}/sitemap.xml`, sitemap(new Date().toISOString().slice(0, 10)));
}
