import { expect, test } from 'bun:test';
import { CANVAS_PAGE_GAP_PX } from '@betteroffice/docx/layout/render';
import type { RetainedFrame } from '@betteroffice/docx/layout/render';
import { commandPageGeometry } from './commandPageGeometry';

function frame(count: number): RetainedFrame {
  return { pages: Array.from({ length: count }, (_, pageIndex) => ({
    pageId: BigInt(pageIndex + 1), pageIndex, primitiveIds: [],
    page: { pageIndex, width: pageIndex < 3 ? 600 : 900, height: 800, primitives: [], unbuilt: pageIndex >= 3 },
  })) } as unknown as RetainedFrame;
}

test('certified prefix grows to full page geometry without moving visible pages', () => {
  const prefix = commandPageGeometry(frame(3), { finalizedPages: 3, documentWidth: 900 });
  const completed = commandPageGeometry(frame(80), { documentWidth: 900 });
  expect(prefix).toHaveLength(3);
  expect(completed).toHaveLength(80);
  expect(prefix.map(page => page.bounds)).toEqual(completed.slice(0, 3).map(page => page.bounds));
  expect(prefix[0]!.bounds.x).toBe(150);
  expect(completed[79]!.bounds.y).toBe(79 * (800 + CANVAS_PAGE_GAP_PX));
});

test('uncertified shells never increase scroll height or acquire geometry', () => {
  const prefix = commandPageGeometry(frame(32), { finalizedPages: 3, documentWidth: 900 });
  expect(prefix.map(page => page.page.pageIndex)).toEqual([0, 1, 2]);
  expect(prefix.at(-1)!.bounds.y + prefix.at(-1)!.bounds.height).toBe(2400 + 2 * CANVAS_PAGE_GAP_PX);
});
