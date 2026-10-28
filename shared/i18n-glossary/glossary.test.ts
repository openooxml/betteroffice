import { expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

const root = resolve(import.meta.dir, '../..');
const PACKAGES = ['docx-i18n', 'xlsx-i18n', 'pptx-i18n'];
const terms = JSON.parse(readFileSync(join(import.meta.dir, 'terms.json'), 'utf8')) as Record<string, string>;

function readJson(path: string): Record<string, unknown> {
  return JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
}

function leafEntries(value: unknown, prefix = ''): Array<[string, unknown]> {
  if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
    return Object.entries(value as Record<string, unknown>)
      .filter(([key]) => prefix !== '' || key !== '_lang')
      .flatMap(([key, child]) => leafEntries(child, prefix ? `${prefix}.${key}` : key));
  }
  return [[prefix, value]];
}

function normalize(text: string, code: string): string {
  return text.trim().replace(/(\.\.\.|…|:)$/u, '').trim().toLocaleLowerCase(code);
}

const codes = [...new Set(PACKAGES.flatMap((name) => readdirSync(join(root, 'packages', name))))]
  .filter((file) => /^[a-z]{2,3}(-[a-zA-Z0-9]{2,8})*\.json$/.test(file))
  .map((file) => file.replace(/\.json$/, ''))
  .filter((code) => code !== 'en')
  .sort();

for (const code of codes) {
  test(`${code} glossary covers every term`, () => {
    const path = join(import.meta.dir, `${code}.json`);
    expect(existsSync(path)).toBe(true);
    const glossary = readJson(path);
    const empty = Object.entries(glossary)
      .filter(([, value]) => ![value].flat().every((entry) => typeof entry === 'string' && entry.trim() !== ''))
      .map(([term]) => term);
    expect({ keys: Object.keys(glossary).sort(), empty }).toEqual({ keys: Object.keys(terms).sort(), empty: [] });
  });

  test(`${code} labels use the glossary terms`, () => {
    const path = join(import.meta.dir, `${code}.json`);
    if (!existsSync(path)) return;
    const glossary = readJson(path);
    const accepted = new Map(
      Object.entries(glossary).map(([term, value]) => [
        normalize(term, 'en'),
        new Set([value].flat().map((entry) => normalize(String(entry), code))),
      ]),
    );
    const mismatched: string[] = [];
    for (const name of PACKAGES) {
      const localePath = join(root, 'packages', name, `${code}.json`);
      if (!existsSync(localePath)) continue;
      const translated = new Map(leafEntries(readJson(localePath)));
      for (const [key, source] of leafEntries(readJson(join(root, 'packages', name, 'en.json')))) {
        if (typeof source !== 'string') continue;
        const allowed = accepted.get(normalize(source, 'en'));
        const value = translated.get(key);
        if (!allowed || typeof value !== 'string') continue;
        if (!allowed.has(normalize(value, code))) mismatched.push(`${name} ${key}: "${value}"`);
      }
    }
    expect(mismatched).toEqual([]);
  });
}
