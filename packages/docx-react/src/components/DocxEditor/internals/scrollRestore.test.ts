import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type {
  DisplayListQueries,
  DisplayListRect,
  DisplayListVisualLine,
} from '@betteroffice/docx/layout/render';
import {
  createYrsInputPositionMap,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  createYrsSession,
  type YrsSession,
  type YrsStickyPosition,
} from '@betteroffice/docx/yrs';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

import {
  captureDisplayListScrollAnchor,
  captureDisplayListViewportAnchor,
  restoreDisplayListScrollAnchor,
  restoreDisplayListViewportAnchor,
} from './scrollRestore';
import { PendingScrollRestoreController } from './viewportAnchoring';
import { YrsPositionProjection } from './yrsPositionProjection';

const WASM = resolve(
  import.meta.dir,
  '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
);

// `resolveDisplayPageClientRect`'s canvas-free projection defaults.
const PAGE_HEIGHT = 800;
const PAGE_GAP = 24;
const PAGE_PADDING = 24;
const SCROLLER_TOP = 60;
const SCROLLER_HEIGHT = 200;

/** Column-space top of a page-local y, the space `scrollTop` moves through. */
function documentTop(pageIndex: number, pageY: number): number {
  return PAGE_PADDING + pageIndex * (PAGE_HEIGHT + PAGE_GAP) + pageY;
}

function domRect(top: number, left: number, width: number, height: number): DOMRect {
  return {
    x: left,
    y: top,
    top,
    left,
    right: left + width,
    bottom: top + height,
    width,
    height,
    toJSON: () => ({}),
  };
}

function pageHost(): HTMLElement {
  return {
    querySelector: () => null,
    getBoundingClientRect: () => domRect(0, 0, 600, 1_000),
    offsetWidth: 600,
    clientWidth: 600,
  } as unknown as HTMLElement;
}

function scrollParent(scrollTop = 600): HTMLElement {
  const style = {
    overflowAnchor: '',
    setProperty(name: string, value: string) {
      if (name === 'overflow-anchor') this.overflowAnchor = value;
    },
  };
  return {
    style,
    scrollTop,
    scrollHeight: 3_000,
    clientHeight: SCROLLER_HEIGHT,
    getBoundingClientRect: () => domRect(SCROLLER_TOP, 0, 600, SCROLLER_HEIGHT),
  } as unknown as HTMLElement;
}

/**
 * Host + scroller pair whose geometry tracks a live `scrollTop` and page count,
 * so a page added or removed above the viewport moves client rects exactly the
 * way the browser would.
 */
function createScene(initialScrollTop: number, initialPageCount: number) {
  const state = { scrollTop: initialScrollTop, pageCount: initialPageCount };
  const contentHeight = (): number =>
    PAGE_PADDING * 2 +
    state.pageCount * PAGE_HEIGHT +
    Math.max(0, state.pageCount - 1) * PAGE_GAP;
  const host = {
    querySelector: () => null,
    getBoundingClientRect: () =>
      domRect(SCROLLER_TOP - state.scrollTop, 0, 600, contentHeight()),
    offsetWidth: 600,
    clientWidth: 600,
  } as unknown as HTMLElement;
  const style = {
    overflowAnchor: '',
    setProperty(name: string, value: string) {
      if (name === 'overflow-anchor') this.overflowAnchor = value;
    },
  };
  const scroller = {
    style,
    get scrollTop() {
      return state.scrollTop;
    },
    set scrollTop(next: number) {
      state.scrollTop = Math.min(
        Math.max(0, next),
        Math.max(0, contentHeight() - SCROLLER_HEIGHT)
      );
    },
    get scrollHeight() {
      return contentHeight();
    },
    clientHeight: SCROLLER_HEIGHT,
    getBoundingClientRect: () => domRect(SCROLLER_TOP, 0, 600, SCROLLER_HEIGHT),
  } as unknown as HTMLElement;
  return { state, host, scroller };
}

function visualLine(
  paraId: string,
  from: number,
  to: number,
  y: number,
  pageIndex = 0
): DisplayListVisualLine {
  return {
    pageIndex,
    x: 0,
    y,
    width: 500,
    height: 16,
    baseline: y + 12,
    from,
    to,
    paraId,
  };
}

function queries(
  lines: readonly DisplayListVisualLine[],
  pageCount = 1,
  anchorRect?: DisplayListRect | null
): DisplayListQueries {
  return {
    pageCount: () => pageCount,
    pageSize: () => ({ width: 600, height: PAGE_HEIGHT }),
    visualLines: () => lines,
    anchorRect: () => anchorRect ?? null,
  } as unknown as DisplayListQueries;
}

function caretRect(pageIndex: number, y: number): DisplayListRect {
  return { pageIndex, x: 0, y, width: 2, height: 16 };
}

function inputMap(session: YrsSession, story = 'body') {
  return createYrsInputPositionMap(
    story,
    session.paragraphs(story).map((paragraph) => ({
      paraId: paragraph.paraId,
      length: paragraph.text.length,
    }))
  );
}

const STICKY: YrsStickyPosition = { story: 'body', encoded: Uint8Array.of(1) };

describe('display-list viewport restore integration', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  test('cancels an async restore and its follow-up after local navigation', async () => {
    const paraId = 'paragraph';
    const before = queries([
      visualLine(paraId, 1, 20, 0),
      visualLine(paraId, 21, 40, 40),
    ]);
    const after = queries([
      visualLine(paraId, 1, 20, 0),
      visualLine(paraId, 21, 40, 140),
    ]);
    const host = pageHost();
    const scroller = scrollParent();
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);
    const controller = new PendingScrollRestoreController<typeof anchor>();
    const ticket = controller.capture(anchor);
    const commit = Promise.resolve().then(() => {
      const pending = controller.take();
      if (!pending) return false;
      return controller.run(pending, () =>
        restoreDisplayListViewportAnchor(pending.value, after, host, scroller, () => 21)
      );
    });

    scroller.scrollTop = 1_100;
    controller.cancel();

    expect(await commit).toBe(false);
    expect(scroller.scrollTop).toBe(1_100);
    expect(
      controller.run(ticket, () =>
        restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 21)
      )
    ).toBe(false);
    expect(scroller.scrollTop).toBe(1_100);
  });

  test('cancels the next-frame application after navigation follows the commit', () => {
    const paraId = 'paragraph';
    const before = queries([
      visualLine(paraId, 1, 20, 0),
      visualLine(paraId, 21, 40, 40),
    ]);
    const after = queries([
      visualLine(paraId, 1, 20, 0),
      visualLine(paraId, 21, 40, 140),
    ]);
    const host = pageHost();
    const scroller = scrollParent();
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);
    const controller = new PendingScrollRestoreController<typeof anchor>();
    const ticket = controller.capture(anchor);

    expect(
      controller.run(ticket, () =>
        restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 21)
      )
    ).toBe(true);
    expect(scroller.scrollTop).toBe(700);

    scroller.scrollTop = 1_200;
    controller.cancel();

    expect(
      controller.run(ticket, () =>
        restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 21)
      )
    ).toBe(false);
    expect(scroller.scrollTop).toBe(1_200);
  });

  test('resolves a sticky line through a remote insert in the same paragraph', async () => {
    const local = await createYrsSession({ clientId: 70_001 });
    const remote = await createYrsSession({ clientId: 70_002 });
    try {
      const receipt = local.loadStories([
        {
          storyId: 'body',
          paragraphs: [{ text: 'abcdefghijklmnopqrstuvwxyz'.repeat(4) }],
        },
      ]);
      remote.loadState(local.encodeState());
      const paraId = receipt.body[0];
      const host = pageHost();
      const scroller = scrollParent();
      const before = queries([
        visualLine(paraId, 1, 20, 0),
        visualLine(paraId, 21, 40, 40),
      ]);
      const anchor = captureDisplayListViewportAnchor(
        before,
        host,
        scroller,
        (displayPosition) => {
          const loc = displayPositionToYrsLoc(inputMap(local), displayPosition);
          return loc ? local.encodeStickyPosition(loc) : null;
        }
      );
      const vector = local.encodeStateVector();

      remote.insertText({ story: 'body', paraId, offset: 0 }, '12345');
      local.applyUpdate(remote.encodeStateAsUpdate(vector));

      const resolved =
        anchor.target?.kind === 'position'
          ? local.resolveStickyPosition(anchor.target.position)
          : null;
      expect(resolved).toEqual({ story: 'body', paraId, offset: 25 });

      const after = queries([
        visualLine(paraId, 1, 10, 0),
        visualLine(paraId, 11, 25, 80),
        visualLine(paraId, 26, 45, 140),
      ]);
      restoreDisplayListViewportAnchor(anchor, after, host, scroller, (position) => {
        const loc = local.resolveStickyPosition(position);
        return loc ? yrsLocToDisplayPosition(inputMap(local), loc) : null;
      });

      expect(scroller.scrollTop).toBe(700);
    } finally {
      local.destroy();
      remote.destroy();
    }
  });

  test('projects a table-cell anchor through a remote edit above it', async () => {
    const local = await createYrsSession({ clientId: 70_003 });
    const remote = await createYrsSession({ clientId: 70_004 });
    try {
      const receipt = local.loadStories([
        {
          storyId: 'body',
          paragraphs: [{ text: 'before' }, { text: 'after' }],
        },
      ]);
      const table = local.insertTable(
        { story: 'body', paraId: receipt.body[1], offset: 0 },
        1,
        1
      );
      const cellStory = table.createdStoryIds[0];
      const cellParaId = local.paragraphs(cellStory)[0].paraId;
      local.insertText({ story: cellStory, paraId: cellParaId, offset: 0 }, 'cell anchor');
      remote.loadState(local.encodeState());

      const beforeProjection = new YrsPositionProjection(local, 'body');
      const beforePosition = beforeProjection.positionForLoc({
        story: cellStory,
        paraId: cellParaId,
        offset: 0,
      });
      expect(beforePosition).not.toBeNull();

      const host = pageHost();
      const scroller = scrollParent();
      const anchor = captureDisplayListViewportAnchor(
        queries([visualLine(cellParaId, beforePosition!, beforePosition! + 11, 60)]),
        host,
        scroller,
        (displayPosition) => {
          const target = beforeProjection.targetAt(displayPosition);
          const loc = displayPositionToYrsLoc(
            inputMap(local, target.story),
            target.displayPosition
          );
          return loc ? local.encodeStickyPosition(loc) : null;
        }
      );
      const vector = local.encodeStateVector();

      remote.insertText({ story: 'body', paraId: receipt.body[0], offset: 0 }, 'remote ');
      local.applyUpdate(remote.encodeStateAsUpdate(vector));

      const afterProjection = new YrsPositionProjection(local, 'body');
      const afterPosition = afterProjection.positionForLoc({
        story: cellStory,
        paraId: cellParaId,
        offset: 0,
      });
      expect(afterPosition).toBe(beforePosition! + 7);

      restoreDisplayListViewportAnchor(
        anchor,
        queries([visualLine(cellParaId, afterPosition!, afterPosition! + 11, 160)]),
        host,
        scroller,
        (position) => {
          const loc = local.resolveStickyPosition(position);
          return loc ? afterProjection.positionForLoc(loc) : null;
        }
      );

      expect(scroller.scrollTop).toBe(700);
    } finally {
      local.destroy();
      remote.destroy();
    }
  });

  test('falls back to the captured scroll position when the sticky line is lost', () => {
    const paraId = 'paragraph';
    const geometry = queries([
      visualLine(paraId, 1, 20, 0),
      visualLine(paraId, 21, 40, 40),
    ]);
    const host = pageHost();
    const scroller = scrollParent();
    const anchor = captureDisplayListViewportAnchor(geometry, host, scroller, () => STICKY);

    scroller.scrollTop = 900;
    restoreDisplayListViewportAnchor(anchor, geometry, host, scroller, () => null);

    expect(scroller.scrollTop).toBe(600);
  });
});

describe('viewport anchoring across a page boundary', () => {
  const paraId = 'paragraph';

  test('compensates when the anchored line reflows onto the next page', () => {
    // Anchored line sits 24px into the viewport near the foot of page 0.
    const { host, scroller } = createScene(700, 2);
    const before = queries([visualLine(paraId, 1, 20, 700, 0)], 2);
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);
    expect(anchor.target?.kind).toBe('position');
    expect(anchor.viewportOffset).toBe(documentTop(0, 700) - 700);

    // Overflow pushes the same line onto page 1.
    const after = queries([visualLine(paraId, 1, 20, 100, 1)], 2);
    restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 1);

    expect(scroller.scrollTop).toBe(documentTop(1, 100) - anchor.viewportOffset);
    expect(scroller.scrollTop).not.toBe(anchor.scrollTopSnapshot);
  });

  test('compensates when a page is added above the viewport', () => {
    const { state, host, scroller } = createScene(1_000, 2);
    const before = queries([visualLine(paraId, 1, 20, 200, 1)], 2);
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);
    const anchoredClientTop = SCROLLER_TOP - 1_000 + documentTop(1, 200);
    expect(anchor.viewportOffset).toBe(anchoredClientTop - SCROLLER_TOP);

    state.pageCount = 3;
    const after = queries([visualLine(paraId, 1, 20, 200, 2)], 3);
    restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 1);

    // Same content, same place on screen.
    expect(scroller.scrollTop).toBe(documentTop(2, 200) - anchor.viewportOffset);
    expect(SCROLLER_TOP - scroller.scrollTop + documentTop(2, 200)).toBe(anchoredClientTop);
  });

  test('compensates when a page is removed above the viewport', () => {
    const { state, host, scroller } = createScene(1_824, 3);
    const before = queries([visualLine(paraId, 1, 20, 200, 2)], 3);
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);
    const anchoredClientTop = SCROLLER_TOP - 1_824 + documentTop(2, 200);

    state.pageCount = 2;
    const after = queries([visualLine(paraId, 1, 20, 200, 1)], 2);
    restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 1);

    expect(scroller.scrollTop).toBe(documentTop(1, 200) - anchor.viewportOffset);
    expect(SCROLLER_TOP - scroller.scrollTop + documentTop(1, 200)).toBe(anchoredClientTop);
  });

  test('does not move when the page boundary reflow is below the viewport', () => {
    const { state, host, scroller } = createScene(1_000, 2);
    const before = queries(
      [visualLine(paraId, 1, 20, 200, 1), visualLine('tail', 40, 60, 700, 1)],
      2
    );
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);

    // The tail paragraph overflows onto a new page; the anchored line does not move.
    state.pageCount = 3;
    const after = queries(
      [visualLine(paraId, 1, 20, 200, 1), visualLine('tail', 40, 60, 100, 2)],
      3
    );
    restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 1);

    expect(scroller.scrollTop).toBe(1_000);
  });

  test('anchors to the nearest line when the viewport top sits in a page gap', () => {
    // Viewport top lands between page 0 and page 1, so no line intersects it.
    const gapScrollTop = documentTop(0, PAGE_HEIGHT) + 8;
    const { state, host, scroller } = createScene(gapScrollTop, 2);
    const before = queries([visualLine(paraId, 1, 20, 700, 0)], 2);
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);
    expect(anchor.target?.kind).toBe('position');

    state.pageCount = 3;
    const after = queries([visualLine(paraId, 1, 20, 700, 1)], 3);
    restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => 1);

    expect(scroller.scrollTop).toBe(documentTop(1, 700) - anchor.viewportOffset);
  });

  test('holds the captured offset when the anchor is lost at a page boundary', () => {
    const { state, host, scroller } = createScene(1_000, 2);
    const before = queries([visualLine(paraId, 1, 20, 200, 1)], 2);
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => STICKY);

    state.pageCount = 3;
    const after = queries([], 3);
    restoreDisplayListViewportAnchor(anchor, after, host, scroller, () => null);

    expect(scroller.scrollTop).toBe(1_000);
  });
});

describe('local selection anchoring across a page boundary', () => {
  test('pins the caret line while it stays on its page', () => {
    const { host, scroller } = createScene(700, 2);
    const before = queries([], 2, caretRect(0, 700));
    const anchor = captureDisplayListScrollAnchor(before, host, scroller, 42);
    expect(anchor.pageIndex).toBe(0);

    // The caret line slid 40px down its own page (a paragraph above grew).
    const after = queries([], 2, caretRect(0, 740));
    restoreDisplayListScrollAnchor(anchor, after, host, scroller);

    expect(scroller.scrollTop).toBe(740);
  });

  test('holds the captured offset when the caret line reflows onto the next page', () => {
    const { host, scroller } = createScene(700, 2);
    const before = queries([], 2, caretRect(0, 700));
    const anchor = captureDisplayListScrollAnchor(before, host, scroller, 42);

    const after = queries([], 2, caretRect(1, 100));
    restoreDisplayListScrollAnchor(anchor, after, host, scroller);

    // Pinning would have dragged the viewport a whole page break down.
    expect(scroller.scrollTop).toBe(700);
  });

  test('holds the captured offset when the caret rect no longer projects', () => {
    const { state, host, scroller } = createScene(700, 2);
    const before = queries([], 2, caretRect(0, 700));
    const anchor = captureDisplayListScrollAnchor(before, host, scroller, 42);

    // A page-count change moves `maxScroll`; the hold must ignore it.
    state.pageCount = 3;
    const after = queries([], 3, null);
    restoreDisplayListScrollAnchor(anchor, after, host, scroller);

    expect(scroller.scrollTop).toBe(700);
  });
});
