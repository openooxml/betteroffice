import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
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

const PLURAL = /\{(\w+),\s*plural,((?:[^{}]|\{[^{}]*\})*)\}/g;

function placeholders(text: string): string[] {
  const names = [
    ...[...text.matchAll(PLURAL)].map((match) => match[1]),
    ...[...text.replace(PLURAL, '').matchAll(/\{(\w+)\}/g)].map((match) => match[1]),
  ];
  return [...new Set(names)].sort();
}

function pluralCategories(text: string): string[][] {
  return [...text.matchAll(PLURAL)].map((match) =>
    [...match[2].matchAll(/(=\d+|\w+)\s*\{/g)].map((branch) => branch[1]),
  );
}

const enEntries = leafEntries(en);
const expectedKeys = new Set(enEntries.map(([key]) => key));
const expectedPlaceholders = new Map(
  enEntries
    .filter((entry): entry is [string, string] => typeof entry[1] === 'string')
    .map(([key, text]) => [key, placeholders(text)] as [string, string[]]),
);

function expectNone(code: string, problem: string, items: string[]) {
  if (items.length === 0) return;
  throw new Error(
    `packages/${basename(packageDir)}/${code}.json ${problem}:\n  ${items.join('\n  ')}\nAdd the translation to every locale; see CONTRIBUTING.md#translations.`,
  );
}

const otherCodes = localeCodes.filter((code) => code !== 'en');
const dataByCode = new Map(
  otherCodes.map((code) => [
    code,
    JSON.parse(readFileSync(join(packageDir, `${code}.json`), 'utf8')) as Record<string, unknown>,
  ]),
);

test('locale files on disk match the exported locales', () => {
  expect(localeCodes.length).toBeGreaterThan(0);
  expect(Object.keys(locales).sort()).toEqual(localeCodes);
});

for (const code of otherCodes) {
  const data = dataByCode.get(code);

  test(`${code} declares its own _lang`, () => {
    expect(data?._lang).toBe(code);
  });

  test(`${code} matches the en key set`, () => {
    const actual = new Set(leafEntries(data).map(([key]) => key));
    const missing = [...expectedKeys].filter((key) => !actual.has(key));
    const extra = [...actual].filter((key) => !expectedKeys.has(key));
    expectNone(code, 'is missing keys that en.json has', missing);
    expectNone(code, 'has keys that en.json does not have', extra);
  });

  test(`${code} translates every key`, () => {
    const untranslated = leafEntries(data)
      .filter(([, text]) => typeof text !== 'string' || text.trim() === '')
      .map(([key]) => key);
    expectNone(code, 'has untranslated keys', untranslated);
  });

  test(`${code} keeps the en placeholders`, () => {
    const mismatched: string[] = [];
    for (const [key, text] of leafEntries(data)) {
      if (typeof text !== 'string') continue;
      const expected = expectedPlaceholders.get(key);
      if (expected === undefined) continue;
      const actual = placeholders(text);
      if (actual.join('\0') !== expected.join('\0')) {
        mismatched.push(`${key}: expected {${expected.join('}, {')}}, found {${actual.join('}, {')}}`);
      }
    }
    expectNone(code, 'changes placeholders', mismatched);
  });

  test(`${code} plurals cover every CLDR category`, () => {
    const supported = new Intl.PluralRules(code).resolvedOptions().pluralCategories;
    const wrong: string[] = [];
    for (const [key, text] of leafEntries(data)) {
      if (typeof text !== 'string') continue;
      for (const branches of pluralCategories(text)) {
        const named = branches.filter((branch) => !branch.startsWith('='));
        const missing = supported.filter((category) => !named.includes(category));
        const unknown = named.filter((category) => !supported.includes(category as Intl.LDMLPluralRule));
        if (missing.length > 0 || unknown.length > 0) {
          wrong.push(`${key}: plural needs exactly ${supported.join(', ')} (missing ${missing.join(', ') || 'none'}, unknown ${unknown.join(', ') || 'none'})`);
        }
      }
    }
    expectNone(code, 'has plurals without the CLDR categories', wrong);
  });
}
