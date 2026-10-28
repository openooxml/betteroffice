import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { HeaderFooterSettings, SlideHeaderFooterText, SlideDateText, SlideHeaderFooter, PptxExportList, SlideLayoutSnapshot } from './index';

const index = readFileSync(resolve(import.meta.dir, 'index.ts'), 'utf8');

test('read fact types are exported from the root', () => {
  const list: PptxExportList = { kind: 'number', scheme: 'arabicPeriod', startAt: 1, value: 1, marker: '1.' };
  const layout: SlideLayoutSnapshot = { name: 'Heading', type: 'title' };
  expect([list.marker, layout.type]).toEqual(['1.', 'title']);
  expect(index).toMatch(/export type\s*\{[^}]*\bPptxExportList\b[^}]*\bSlideLayoutSnapshot\b[^}]*\}\s*from\s*['"]@betteroffice\/pptx['"]/);
});

test('header footer read types are exported from the root', () => {
  const settings: HeaderFooterSettings = { slideNumber: true, footer: false, date: true, header: false };
  const footer: SlideHeaderFooterText = { shapeId: 'shape', text: '' };
  const date: SlideDateText = { shapeId: 'date', text: 'Fixed date', field: null };
  const facts: SlideHeaderFooter = { footer, slideNumber: null, date };
  const layout: SlideLayoutSnapshot = { name: null, type: null, headerFooter: settings };
  expect([facts.footer?.text, facts.date?.field, layout.headerFooter?.slideNumber]).toEqual(['', null, true]);
  for (const name of ['HeaderFooterSettings', 'SlideHeaderFooterText', 'SlideDateText', 'SlideHeaderFooter']) {
    expect(index).toMatch(new RegExp(`export type\\s*\\{[^}]*\\b${name}\\b[^}]*\\}\\s*from\\s*['"]@betteroffice/pptx['"]`));
  }
});
