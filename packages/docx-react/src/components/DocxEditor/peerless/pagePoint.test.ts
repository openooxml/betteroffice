import { expect, test } from 'bun:test';
import { clientPointToPage } from './pagePoint';
import type { PeerlessFrameGeometry } from './viewStore';

test('client points use page-local layout coordinates at 150 percent on offset stacked pages', () => {
  const pages: PeerlessFrameGeometry['pages'] = [
    { page: { pageId: '11', pageIndex: 0 }, bounds: { x: 20, y: 0, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] },
    { page: { pageId: '22', pageIndex: 1 }, bounds: { x: 60, y: 840, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] },
  ];
  const host = { left: 40, top: -800, width: 990, height: 2460 };
  const pointer = clientPointToPage(310, 514, host, 1.5, pages);
  const read = clientPointToPage(310, 514, { ...host }, 1.5, pages);
  expect(pointer).toEqual({ page: pages[1], x: 120, y: 36 });
  expect(read).toEqual(pointer);
});
