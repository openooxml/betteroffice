import { expect, test } from 'bun:test';
import { bodyEdgeY, clientPointToPage } from './pagePoint';
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

test('band presses map inside the body, past headers, footers and note areas that reach into it', () => {
  const contentBounds = { x: 96, y: 288, width: 624, height: 480 };
  expect(bodyEdgeY({ contentBounds }, 'header')).toBe(289);
  expect(bodyEdgeY({ contentBounds }, 'footer')).toBe(767);
  const notes = { contentBounds, noteAreas: [{ y: 700, height: 68 }, { y: 650, height: 51 }] };
  expect(bodyEdgeY(notes, 'footer')).toBe(649);
  expect(bodyEdgeY(notes, 'header')).toBe(289);
  const tallHeader = { contentBounds, header: { rId: 'h', kind: 'header' as const, y: 48, height: 260, primitives: [] } };
  expect(bodyEdgeY(tallHeader, 'header')).toBe(309);
  expect(bodyEdgeY({ contentBounds, noteAreas: [{ y: 288, height: 480 }] }, 'footer')).toBeNull();
  expect(bodyEdgeY({}, 'header')).toBeNull();
  expect(bodyEdgeY({ contentBounds, unbuilt: true }, 'footer')).toBeNull();
});
