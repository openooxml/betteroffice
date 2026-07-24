import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type {
  DisplayListQueries,
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
  captureDisplayListViewportAnchor,
  restoreDisplayListViewportAnchor,
} from './scrollRestore';
import { PendingScrollRestoreController } from './viewportAnchoring';
import { YrsPositionProjection } from './yrsPositionProjection';

const WASM = resolve(
  import.meta.dir,
  '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
);

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
    clientHeight: 200,
    getBoundingClientRect: () => domRect(60, 0, 600, 200),
  } as unknown as HTMLElement;
}

function visualLine(
  paraId: string,
  from: number,
  to: number,
  y: number
): DisplayListVisualLine {
  return {
    pageIndex: 0,
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

function queries(lines: readonly DisplayListVisualLine[]): DisplayListQueries {
  return {
    pageCount: () => 1,
    pageSize: () => ({ width: 600, height: 800 }),
    visualLines: () => lines,
  } as unknown as DisplayListQueries;
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
    const sticky: YrsStickyPosition = {
      story: 'body',
      encoded: Uint8Array.of(1),
    };
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => sticky);
    const controller = new PendingScrollRestoreController<typeof anchor>();
    const ticket = controller.capture(anchor);
    const commit = Promise.resolve().then(() => {
      const pending = controller.take();
      if (!pending) return false;
      return controller.run(pending, () =>
        restoreDisplayListViewportAnchor(
          pending.value,
          after,
          host,
          scroller,
          () => 21
        )
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
    const anchor = captureDisplayListViewportAnchor(before, host, scroller, () => ({
      story: 'body',
      encoded: Uint8Array.of(1),
    }));
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
        (displayPosition, expectedParaId) => {
          const loc = displayPositionToYrsLoc(inputMap(local), displayPosition);
          return loc?.paraId === expectedParaId ? local.encodeStickyPosition(loc) : null;
        }
      );
      const vector = local.encodeStateVector();

      remote.insertText({ story: 'body', paraId, offset: 0 }, '12345');
      local.applyUpdate(remote.encodeStateAsUpdate(vector));

      const resolved =
        anchor.target?.kind === 'paragraph'
          ? local.resolveStickyPosition(anchor.target.position)
          : null;
      expect(resolved).toEqual({ story: 'body', paraId, offset: 25 });

      const after = queries([
        visualLine(paraId, 1, 10, 0),
        visualLine(paraId, 11, 25, 80),
        visualLine(paraId, 26, 45, 140),
      ]);
      restoreDisplayListViewportAnchor(
        anchor,
        after,
        host,
        scroller,
        (position, expectedParaId) => {
          const loc = local.resolveStickyPosition(position);
          return loc?.paraId === expectedParaId
            ? yrsLocToDisplayPosition(inputMap(local), loc)
            : null;
        }
      );

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
        (displayPosition, expectedParaId) => {
          const target = beforeProjection.targetAt(displayPosition);
          const loc = displayPositionToYrsLoc(inputMap(local, target.story), target.displayPosition);
          return loc?.paraId === expectedParaId ? local.encodeStickyPosition(loc) : null;
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
        (position, expectedParaId) => {
          const loc = local.resolveStickyPosition(position);
          return loc?.paraId === expectedParaId ? afterProjection.positionForLoc(loc) : null;
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
    const anchor = captureDisplayListViewportAnchor(geometry, host, scroller, () => ({
      story: 'body',
      encoded: Uint8Array.of(1),
    }));

    scroller.scrollTop = 900;
    restoreDisplayListViewportAnchor(anchor, geometry, host, scroller, () => null);

    expect(scroller.scrollTop).toBe(600);
  });
});
