import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { PptxExportList, SlideLayoutSnapshot } from './index';

const index = readFileSync(resolve(import.meta.dir, 'index.ts'), 'utf8');

test('read fact types are exported from the root', () => {
  const list: PptxExportList = { kind: 'number', scheme: 'arabicPeriod', startAt: 1, value: 1, marker: '1.' };
  const layout: SlideLayoutSnapshot = { name: 'Heading', type: 'title' };
  expect([list.marker, layout.type]).toEqual(['1.', 'title']);
  expect(index).toMatch(/export type\s*\{[^}]*\bPptxExportList\b[^}]*\bSlideLayoutSnapshot\b[^}]*\}\s*from\s*['"]@betteroffice\/pptx['"]/);
});
