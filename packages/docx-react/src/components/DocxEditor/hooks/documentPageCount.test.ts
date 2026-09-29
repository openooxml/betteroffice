import { expect, test } from 'bun:test';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { documentPageCount } from './documentPageCount';

test('a partial layout reports no page count', () => {
  const layout = { pageSize: { w: 1, h: 1 }, pages: [{}, {}] } as unknown as Layout;
  expect(documentPageCount(layout)).toBe(2);
  expect(documentPageCount({ ...layout, partial: true })).toBe(0);
  expect(documentPageCount(null)).toBe(0);
});
