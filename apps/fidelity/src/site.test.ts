import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { PAGES, robots, sitemap } from './site';

test('robots allows everything and names the sitemap', () => {
  expect(robots()).toBe('User-agent: *\nAllow: /\n\nSitemap: https://benchmarks.betteroffice.dev/sitemap.xml\n');
});

test('the sitemap lists every public page with a lastmod', () => {
  const xml = sitemap('2026-10-27');
  expect(xml).toContain('<loc>https://benchmarks.betteroffice.dev/</loc><lastmod>2026-10-27</lastmod>');
  expect(xml).toContain('<loc>https://benchmarks.betteroffice.dev/compare</loc><lastmod>2026-10-27</lastmod>');
  expect(xml.match(/<url>/g)).toHaveLength(PAGES.length);
});

test('pages are indexable and canonical to their own URL', () => {
  for (const [file, path] of [['index.html', '/'], ['compare.html', '/compare']] as const) {
    const html = readFileSync(join(import.meta.dir, file), 'utf8');
    expect(html).not.toMatch(/name="robots"/);
    expect(html).toContain(`<link rel="canonical" href="https://benchmarks.betteroffice.dev${path}" />`);
    expect(html).toMatch(/name="description"/);
  }
});
