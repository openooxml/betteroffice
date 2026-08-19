/**
 * Clicking into a footnote: the pointer path opens the note it lands in, and
 * the caret that click asked for goes into that note's story once the editor
 * is on it. `partEdit.test.ts` covers the mapping; this shows the wiring.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { createRef, useState } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';
import type { YrsInputRef } from '../YrsInput';
import { partEditStory, type PartEdit } from '../partEdit';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
/** y at or past this is the page's footnote area in the fake engine. */
const NOTE_FROM = 900;
const NOTE_A_ID = 1;
const NOTE_ID = 2;
/** what the fake engine resolves anywhere inside the note area */
const NOTE_POSITION = 7;

let host: HTMLDivElement;
let selections: Array<{ anchor: number; head: number; story?: string }>;

function fakeQueries(): DisplayListQueries {
  return {
    displayList: { pages: [{ pageIndex: 0, ...PAGE, primitives: [] }] },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number) => {
      if (y < NOTE_FROM) return { region: 'body' as const, pos: 1, target: 'text' as const };
      return x < PAGE.width / 2
        ? {
            region: 'footnote' as const,
            noteId: NOTE_A_ID,
            pos: 4,
            target: 'text' as const,
          }
        : {
            region: 'footnote' as const,
            noteId: NOTE_ID,
            pos: NOTE_POSITION,
            target: 'text' as const,
          };
    },
    imageAtPoint: () => null,
  } as unknown as DisplayListQueries;
}

function stableOptions(): (overrides?: Partial<UsePagesPointerOptions>) => UsePagesPointerOptions {
  const canvasHostRef = createRef<HTMLDivElement>() as { current: HTMLDivElement | null };
  canvasHostRef.current = host;
  const yrsInputRef = {
    current: {
      focus: () => {},
      setSelectionFromDisplay: (anchor: number, head = anchor, story?: string) => {
        selections.push({ anchor, head, story });
      },
      displaySelection: () => null,
    } as unknown as YrsInputRef,
  };
  // Positions in a root story project back onto that same story, which is what
  // the real projection does for a story with no nested children.
  const projectionFor = (rootStory: string) => {
    const size = 10;
    return {
      size: 10,
      targetAt: (position: number) =>
        position >= 0 && position < size ? { story: rootStory, displayPosition: position } : null,
      tableAtPosition: () => null,
      cellPosition: () => null,
    } as unknown as YrsPositionProjection;
  };
  const base: UsePagesPointerOptions = {
    pagesContainerRef: { current: null },
    yrsInputRef,
    yrsSession: { cellSelection: () => null } as unknown as YrsSession,
    yrsRootStory: 'body',
    getYrsPositionProjection: projectionFor,
    applyYrsCommand: () => false,
    syncYrsInputState: () => false,
    readOnly: false,
    displayListQueries: fakeQueries(),
    canvasHostRef,
    setSelectionRects: () => {},
    setCaretPosition: () => {},
    setIsFocused: () => {},
    scrollToPositionImpl: () => {},
  };
  return (overrides = {}) => ({ ...base, ...overrides });
}

function dispatchMouse(
  type: string,
  clientX: number,
  clientY: number,
  target: EventTarget,
  init: MouseEventInit = {}
): void {
  target.dispatchEvent(
    new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY, button: 0, ...init })
  );
}

function mouse(type: string, clientX: number, clientY: number, target: EventTarget): void {
  act(() => dispatchMouse(type, clientX, clientY, target));
}

beforeEach(() => {
  selections = [];
  host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: PAGE.width,
      bottom: PAGE.height,
      width: PAGE.width,
      height: PAGE.height,
    }) as DOMRect;
  host.append(canvas);
  document.body.append(host);
});

afterEach(() => {
  cleanup();
  host.remove();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const canvasOf = () => host.firstElementChild as HTMLCanvasElement;
const note: PartEdit = { kind: 'footnote', noteId: NOTE_ID };

describe('clicking a note', () => {
  test('a click in the note area opens that note instead of moving the body caret', () => {
    const opened: PartEdit[] = [];
    const options = stableOptions();
    renderHook(() => usePagesPointer(options({ onNoteClick: (hit) => opened.push(hit) })));

    mouse('mousedown', 400, 950, canvasOf());

    expect(opened).toEqual([note]);
    expect(selections).toEqual([]);
  });

  test('the caret the opening click asked for lands in the note story', () => {
    const options = stableOptions();
    renderHook(() => {
      const [partEdit, setPartEdit] = useState<PartEdit | null>(null);
      return usePagesPointer(
        options({
          partEdit,
          yrsRootStory: partEditStory(partEdit),
          onNoteClick: setPartEdit,
        })
      );
    });

    mouse('mousedown', 400, 950, canvasOf());
    expect(selections).toEqual([
      { anchor: NOTE_POSITION, head: NOTE_POSITION, story: `fn:${NOTE_ID}` },
    ]);
  });

  test('the opening caret survives until React commits the requested note', async () => {
    const options = stableOptions();
    const actEnvironment = globalThis as typeof globalThis & {
      IS_REACT_ACT_ENVIRONMENT?: boolean;
    };
    const previousActEnvironment = actEnvironment.IS_REACT_ACT_ENVIRONMENT;
    const { rerender } = renderHook(() => {
      const [partEdit, setPartEdit] = useState<PartEdit | null>(null);
      return usePagesPointer(
        options({
          partEdit,
          yrsRootStory: partEditStory(partEdit),
          onNoteClick: setPartEdit,
        })
      );
    });

    try {
      actEnvironment.IS_REACT_ACT_ENVIRONMENT = false;
      dispatchMouse('mousedown', 400, 950, canvasOf());
      await Promise.resolve();
      actEnvironment.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment;
      act(() => rerender());

      expect(selections).toEqual([
        { anchor: NOTE_POSITION, head: NOTE_POSITION, story: `fn:${NOTE_ID}` },
      ]);
    } finally {
      actEnvironment.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment;
    }
  });

  test('a second note click replaces the first deferred caret', () => {
    const options = stableOptions();
    const noteA = { kind: 'footnote', noteId: NOTE_A_ID } as const;
    const { rerender } = renderHook(
      (partEdit: PartEdit | null) =>
        usePagesPointer(
          options({
            partEdit,
            yrsRootStory: partEditStory(partEdit),
            onNoteClick: () => {},
          })
        ),
      { initialProps: null as PartEdit | null }
    );

    act(() => {
      dispatchMouse('mousedown', 200, 950, canvasOf());
      dispatchMouse('mousedown', 600, 950, canvasOf());
      rerender(note);
    });

    expect(selections).toEqual([
      { anchor: NOTE_POSITION, head: NOTE_POSITION, story: `fn:${NOTE_ID}` },
    ]);
    expect(selections).not.toContainEqual({
      anchor: 4,
      head: 4,
      story: `fn:${noteA.noteId}`,
    });
  });

  test('a later body click cancels the deferred note caret', () => {
    const options = stableOptions();
    const { rerender } = renderHook(
      (partEdit: PartEdit | null) =>
        usePagesPointer(
          options({
            partEdit,
            yrsRootStory: partEdit ? `fn:${NOTE_ID}` : 'body',
            onNoteClick: () => {},
          })
        ),
      { initialProps: null as PartEdit | null }
    );

    act(() => {
      dispatchMouse('mousedown', 600, 950, canvasOf());
      dispatchMouse('mousedown', 400, 500, canvasOf());
      dispatchMouse('mouseup', 400, 500, window);
      rerender(note);
    });

    expect(selections).toEqual([{ anchor: 1, head: 1, story: 'body' }]);
  });

  test('a right-click cancels the deferred note caret', () => {
    const options = stableOptions();
    const { rerender } = renderHook(
      (partEdit: PartEdit | null) =>
        usePagesPointer(
          options({
            partEdit,
            yrsRootStory: partEdit ? `fn:${NOTE_ID}` : 'body',
            onNoteClick: () => {},
          })
        ),
      { initialProps: null as PartEdit | null }
    );

    act(() => {
      dispatchMouse('mousedown', 600, 950, canvasOf());
      dispatchMouse('mousedown', 600, 950, canvasOf(), { button: 2 });
      rerender(note);
    });

    expect(selections).toEqual([]);
  });

  test('a request the host never honours does not leak into a later note open', () => {
    const options = stableOptions();
    const { rerender } = renderHook(
      (partEdit: PartEdit | null) =>
        usePagesPointer(
          options({
            partEdit,
            yrsRootStory: partEdit ? `fn:${NOTE_ID}` : 'body',
            onNoteClick: () => {},
          })
        ),
      { initialProps: null as PartEdit | null }
    );

    mouse('mousedown', 600, 950, canvasOf());
    act(() => rerender(note));

    expect(selections).toEqual([]);
  });

  test('opening a different note discards the deferred note caret', () => {
    const options = stableOptions();
    const otherNote: PartEdit = { kind: 'footnote', noteId: NOTE_A_ID };
    const { rerender } = renderHook(
      ({ partEdit, story }: { partEdit: PartEdit | null; story: string }) =>
        usePagesPointer(options({ partEdit, yrsRootStory: story, onNoteClick: () => {} })),
      { initialProps: { partEdit: null as PartEdit | null, story: 'body' } }
    );

    act(() => {
      dispatchMouse('mousedown', 600, 950, canvasOf());
      rerender({ partEdit: otherNote, story: `fn:${NOTE_A_ID}` });
    });
    act(() => rerender({ partEdit: note, story: `fn:${NOTE_ID}` }));

    expect(selections).toEqual([]);
  });

  test('a replacement session cannot inherit a deferred note caret', () => {
    const options = stableOptions();
    const oldSession = { cellSelection: () => null } as unknown as YrsSession;
    const newSession = { cellSelection: () => null } as unknown as YrsSession;
    const { rerender } = renderHook(
      ({ partEdit, session }: { partEdit: PartEdit | null; session: YrsSession }) =>
        usePagesPointer(
          options({
            partEdit,
            yrsSession: session,
            yrsRootStory: partEdit ? `fn:${NOTE_ID}` : 'body',
            onNoteClick: () => {},
          })
        ),
      {
        initialProps: {
          partEdit: null as PartEdit | null,
          session: oldSession,
        },
      }
    );

    act(() => {
      dispatchMouse('mousedown', 600, 950, canvasOf());
      rerender({ partEdit: note, session: newSession });
    });

    expect(selections).toEqual([]);
  });

  test('revalidates the deferred offset after an intervening edit', () => {
    let size = 10;
    const projectionFor = (rootStory: string) =>
      ({
        size,
        targetAt: (position: number) =>
          position >= 0 && position < size ? { story: rootStory, displayPosition: position } : null,
        tableAtPosition: () => null,
        cellPosition: () => null,
      } as unknown as YrsPositionProjection);
    const options = stableOptions();
    const { rerender } = renderHook(
      (partEdit: PartEdit | null) =>
        usePagesPointer(
          options({
            partEdit,
            yrsRootStory: partEdit ? `fn:${NOTE_ID}` : 'body',
            getYrsPositionProjection: projectionFor,
            onNoteClick: () => {},
          })
        ),
      { initialProps: null as PartEdit | null }
    );

    act(() => {
      dispatchMouse('mousedown', 600, 950, canvasOf());
      size = 4;
      rerender(note);
    });

    expect(selections).toEqual([{ anchor: 3, head: 3, story: `fn:${NOTE_ID}` }]);
  });

  test('a click inside the open note selects in its own story', () => {
    const options = stableOptions();
    renderHook(() =>
      usePagesPointer(options({ partEdit: note, yrsRootStory: `fn:${NOTE_ID}` }))
    );

    mouse('mousedown', 400, 950, canvasOf());

    expect(selections).toEqual([
      { anchor: NOTE_POSITION, head: NOTE_POSITION, story: `fn:${NOTE_ID}` },
    ]);
  });

  test('a click in the body leaves the note instead of typing into it', () => {
    let left = 0;
    const options = stableOptions();
    renderHook(() =>
      usePagesPointer(
        options({ partEdit: note, yrsRootStory: `fn:${NOTE_ID}`, onBodyClick: () => (left += 1) })
      )
    );

    mouse('mousedown', 400, 500, canvasOf());

    expect(left).toBe(1);
    expect(selections).toEqual([]);
  });

  test('a read-only document never opens a note', () => {
    const opened: PartEdit[] = [];
    const options = stableOptions();
    renderHook(() =>
      usePagesPointer(options({ readOnly: true, onNoteClick: (hit) => opened.push(hit) }))
    );

    mouse('mousedown', 400, 950, canvasOf());

    expect(opened).toEqual([]);
  });

  test('a one-unit note image selection has no image context entry', () => {
    const contexts: Array<{ image?: { pos: number } | null }> = [];
    const options = stableOptions();
    const yrsInputRef = {
      current: {
        focus: () => {},
        setSelectionFromDisplay: () => {},
        displaySelection: () => ({ anchor: 3, head: 4 }),
      } as unknown as YrsInputRef,
    };
    const projectionFor = (rootStory: string) =>
      ({
        size: 10,
        targetAt: (position: number) => ({
          story: rootStory,
          displayPosition: position,
        }),
        nodeAt: (position: number) =>
          position === 3 ? { kind: 'image', attrs: { wrapType: 'inline' } } : null,
        tableAtPosition: () => null,
        cellPosition: () => null,
      } as unknown as YrsPositionProjection);
    renderHook(() =>
      usePagesPointer(
        options({
          partEdit: note,
          yrsRootStory: `fn:${NOTE_ID}`,
          yrsInputRef,
          getYrsPositionProjection: projectionFor,
          onContextMenu: (context) => contexts.push(context),
        })
      )
    );

    mouse('contextmenu', 600, 950, canvasOf());

    expect(contexts).toHaveLength(1);
    expect(contexts[0]?.image).toBeNull();
  });
});
