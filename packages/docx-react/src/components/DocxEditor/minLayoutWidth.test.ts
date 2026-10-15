import { expect, test } from 'bun:test';
import { SIDEBAR_DOCUMENT_SHIFT } from '../sidebar/constants';
import { minLayoutWidth } from './minLayoutWidth';

test('zoom 1 reserves the outline allowance on both sides, the page, and the sidebar shift', () => {
  expect(minLayoutWidth({ outlineLeftAllowance: 64, maxPageWidthPx: 1123, zoom: 1, sidebarOpen: false })).toBe(1251);
  expect(minLayoutWidth({ outlineLeftAllowance: 64, maxPageWidthPx: 1123, zoom: 1, sidebarOpen: true })).toBe(
    1251 + SIDEBAR_DOCUMENT_SHIFT * 2
  );
});

test('only the page width scales with zoom', () => {
  expect(minLayoutWidth({ outlineLeftAllowance: 64, maxPageWidthPx: 1123, zoom: 0.73, sidebarOpen: false })).toBe(948);
  expect(minLayoutWidth({ outlineLeftAllowance: 20, maxPageWidthPx: 816, zoom: 2, sidebarOpen: true })).toBe(
    40 + 1632 + SIDEBAR_DOCUMENT_SHIFT * 2
  );
});
