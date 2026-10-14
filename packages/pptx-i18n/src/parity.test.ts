import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { readLocaleCodes } from '../locale-files';
import { locales } from './index';
import en from '../en.json';

const packageDir = resolve(import.meta.dir, '..');
const localeCodes = readLocaleCodes(packageDir);

function leafEntries(value: unknown, prefix = ''): Array<[string, unknown]> {
  if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
    return Object.entries(value as Record<string, unknown>)
      .filter(([key]) => prefix !== '' || key !== '_lang')
      .flatMap(([key, child]) => leafEntries(child, prefix ? `${prefix}.${key}` : key));
  }
  return [[prefix, value]];
}

function placeholders(text: string): string[] {
  const names = [...text.matchAll(/\{(\w+)\}/g)].map((match) => match[1]);
  return [...new Set(names)].sort();
}

const enEntries = leafEntries(en);
const expectedKeys = new Set(enEntries.map(([key]) => key));
const expectedPlaceholders = new Map(
  enEntries
    .filter((entry): entry is [string, string] => typeof entry[1] === 'string')
    .map(([key, text]) => [key, placeholders(text)] as [string, string[]]),
);

const otherCodes = localeCodes.filter((code) => code !== 'en');
const dataByCode = new Map(
  otherCodes.map((code) => [
    code,
    JSON.parse(readFileSync(join(packageDir, `${code}.json`), 'utf8')) as unknown,
  ]),
);

test('locale files on disk cover every exported locale', () => {
  expect(localeCodes.length).toBeGreaterThan(0);
  for (const code of Object.keys(locales)) {
    expect(localeCodes).toContain(code);
  }
});

for (const code of otherCodes) {
  test(`${code} matches the en key set`, () => {
    const actual = new Set(leafEntries(dataByCode.get(code)).map(([key]) => key));
    const missing = [...expectedKeys].filter((key) => !actual.has(key));
    const extra = [...actual].filter((key) => !expectedKeys.has(key));
    expect({ missing, extra }).toEqual({ missing: [], extra: [] });
  });

  test(`${code} keeps the en placeholders`, () => {
    const mismatched: Array<{ key: string; expected: string[]; actual: string[] }> = [];
    for (const [key, text] of leafEntries(dataByCode.get(code))) {
      if (typeof text !== 'string') continue;
      const expected = expectedPlaceholders.get(key);
      if (expected === undefined) continue;
      const actual = placeholders(text);
      if (actual.join('\0') !== expected.join('\0')) {
        mismatched.push({ key, expected, actual });
      }
    }
    expect(mismatched).toEqual([]);
  });
}
