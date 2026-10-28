import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { XlsxSheetVisibility } from './index';

const index = readFileSync(resolve(import.meta.dir, 'index.ts'), 'utf8');

test('sheet visibility type is exported from the root', () => {
  const visibility: XlsxSheetVisibility[] = ['visible', 'hidden', 'veryHidden', 'unknown'];
  expect(visibility).toHaveLength(4);
  expect(index).toMatch(/export\s+type\s*\{[^}]*\bXlsxSheetVisibility\b[^}]*\}\s*from\s*['"]@betteroffice\/xlsx['"]/);
});


test('header footer types are exported from the root', () => {
  const text: import('./index').XlsxHeaderFooterText = { raw: 'Title', left: '', center: 'Title', right: '' };
  const pair: import('./index').XlsxHeaderFooterPair = { header: text, footer: null };
  const facts: import('./index').XlsxSheetHeaderFooter = {
    differentFirst: false, differentOddEven: false, scaleWithDoc: true, alignWithMargins: true,
    odd: pair, even: null, first: null,
  };
  expect(facts.odd.header).toEqual(text);
  for (const name of ['XlsxHeaderFooterText', 'XlsxHeaderFooterPair', 'XlsxSheetHeaderFooter']) {
    expect(index).toMatch(new RegExp(`export\\s+type\\s*\\{[^}]*\\b${name}\\b[^}]*\\}\\s*from\\s*['"]@betteroffice/xlsx['"]`));
  }
});
