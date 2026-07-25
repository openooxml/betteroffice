import { describe, expect, test } from 'bun:test';
import type { DisplayList, DisplayListRect } from '@betteroffice/docx/layout/render';

import {
  buildRemotePresencePageMetrics,
  clampRemoteSelectionRange,
  REMOTE_PRESENCE_MAX_PAGES,
  REMOTE_PRESENCE_MAX_RECTS,
  remotePresencePagePositionRange,
  remotePresencePageWindow,
  simplifyRemoteSelectionRects,
} from './remotePresenceGeometry';

function displayList(pageCount: number): DisplayList {
  return {
    pages: Array.from({ length: pageCount }, (_, pageIndex) => ({
      pageIndex,
      width: 100,
      height: 100,
      primitives: [
        {
          kind: 'text',
          docStart: pageIndex * 100,
          docEnd: (pageIndex + 1) * 100,
        } as DisplayList['pages'][number]['primitives'][number],
      ],
    })),
  };
}

describe('remote presence geometry', () => {
  test('moves a bounded page window with the viewport', () => {
    const metrics = buildRemotePresencePageMetrics(displayList(100), 1);
    const first = remotePresencePageWindow(metrics, 50, 74, 174);
    const scrolled = remotePresencePageWindow(metrics, 50, 4_714, 4_814);
    const oversized = remotePresencePageWindow(metrics, 50, 50, 20_000);

    expect(first).toEqual({ start: 0, end: 1 });
    expect(scrolled).toEqual({ start: 39, end: 41 });
    expect(oversized).not.toBeNull();
    expect((oversized?.end ?? 0) - (oversized?.start ?? 0) + 1).toBe(REMOTE_PRESENCE_MAX_PAGES);
  });

  test('clamps a document-wide selection to the page window', () => {
    const pageRange = remotePresencePagePositionRange(displayList(10), {
      start: 4,
      end: 5,
    });
    expect(pageRange).toEqual({ from: 400, to: 600 });
    if (!pageRange) return;
    expect(clampRemoteSelectionRange(pageRange, 0, 1_000)).toEqual({
      from: 400,
      to: 600,
    });
    expect(clampRemoteSelectionRange(pageRange, 0, 300)).toBeNull();
  });

  test('coalesces selections to a fixed per-peer rect budget', () => {
    const rects: DisplayListRect[] = [];
    for (let pageIndex = 4; pageIndex <= 7; pageIndex += 1) {
      for (let index = 0; index < 200; index += 1) {
        rects.push({
          pageIndex,
          x: index % 2 === 0 ? 10 : 30,
          y: index * 3,
          width: 10,
          height: 2,
        });
      }
    }
    rects.push({ pageIndex: 0, x: 0, y: 0, width: 10, height: 10 });

    const simplified = simplifyRemoteSelectionRects(rects, {
      start: 4,
      end: 7,
    });

    expect(simplified.length).toBeLessThanOrEqual(REMOTE_PRESENCE_MAX_RECTS);
    expect(new Set(simplified.map((rect) => rect.pageIndex))).toEqual(new Set([4, 5, 6, 7]));
    for (let pageIndex = 4; pageIndex <= 7; pageIndex += 1) {
      const pageRects = simplified.filter((rect) => rect.pageIndex === pageIndex);
      expect(Math.min(...pageRects.map((rect) => rect.y))).toBe(0);
      expect(Math.max(...pageRects.map((rect) => rect.y + rect.height))).toBe(599);
    }
  });
});
