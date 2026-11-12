import { readFileSync, readdirSync } from 'node:fs';
import { expect, test } from 'bun:test';

const ENGINE_CRATES = [
  'docx-engine',
  'docx-layout',
  'docx-raster',
  'docx-edit',
  'docx-parse',
  'ooxml-opc',
  'ooxml-text',
  'betteroffice-fonts',
  'ooxml-font-preflight',
];

const SPAWNS = /\brayon\b|thread::(?:spawn|scope|Builder)\b/;

function sources(directory: URL): URL[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = new URL(entry.name + (entry.isDirectory() ? '/' : ''), directory);
    if (entry.isDirectory()) return sources(path);
    return entry.name.endsWith('.rs') ? [path] : [];
  });
}

test.each(ENGINE_CRATES)('%s runs on the calling thread, so a per-thread heap limit sees all its work', (crate) => {
  const root = new URL(`../../crates/${crate}/`, import.meta.url);
  expect(readFileSync(new URL('Cargo.toml', root), 'utf8')).not.toMatch(SPAWNS);
  const spawning = sources(new URL('src/', root))
    .filter((file) => SPAWNS.test(readFileSync(file, 'utf8')))
    .map((file) => file.pathname);
  expect(spawning).toEqual([]);
});
