import { expect, test } from 'bun:test';
import { workerPageAtViewport } from './pageNavigation';
import type { PeerlessFrameGeometry } from './viewStore';

const pages: PeerlessFrameGeometry['pages'] = [0, 1, 2].map((pageIndex) => ({
  page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 124, width: 200, height: 100 },
  transform: [1, 0, 0, 1, 0, 0],
}));

test('page indicator uses worker geometry, zoom and the viewport middle', () => {
  expect(workerPageAtViewport(pages, 150, 60, 1)).toBe(2);
  expect(workerPageAtViewport(pages, 300, 120, 2)).toBe(2);
  expect(workerPageAtViewport(pages, 900, 60, 1)).toBe(3);
});
