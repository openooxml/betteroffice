import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

test('the head-to-head posts the benchmarks page links exist on the site', () => {
  const html = readFileSync(join(import.meta.dir, 'index.html'), 'utf8');
  const slugs = [...html.matchAll(/https:\/\/betteroffice\.dev\/blog\/([a-z0-9-]+)"/g)].map((m) => m[1]);
  expect(slugs).toEqual(['betteroffice-vs-docx-editor', 'betteroffice-vs-univer']);
  for (const slug of slugs) {
    const post = readFileSync(join(import.meta.dir, '..', '..', 'web', 'content', 'blog', `${slug}.md`), 'utf8');
    expect(post).not.toMatch(/^draft: true$/m);
  }
});
