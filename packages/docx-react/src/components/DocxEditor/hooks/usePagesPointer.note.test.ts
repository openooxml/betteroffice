/**
 * Clicking into a footnote: the pointer path opens the note it lands in, and
 * the caret that click asked for goes into that note's story once the editor
 * is on it. `partEdit.test.ts` covers the mapping; this shows the wiring.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { createRef } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';
import type { YrsInputRef } from '../YrsInput';
import type { PartEdit } from '../partEdit';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
/** y at or past this is the page's footnote area in the fake engine. */
const NOTE_FROM = 900;
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
    hitTestRegions: (_pageIndex: number, _x: number, y: number) =>
      y >= NOTE_FROM
        ? { region: 'footnote' as const, noteId: NOTE_ID, pos: NOTE_POSITION, target: 'text' as const }
        : { region: 'body' as const, pos: 1, target: 'text' as const },
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
  const projectionFor = (rootStory: string) =>
    ({
      size: 10,
      targetAt: (position: number) => ({ story: rootStory, displayPosition: position }),
      tableAtPosition: () => null,
      cellPosition: () => null,
    }) as unknown as YrsPositionProjection;
  const base: UsePagesPointerOptions = {
    pagesContainerRef: { current: null },
    yrsInputRef,
    yrsSession: null as unknown as YrsSession,
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

function mouse(type: string, clientX: number, clientY: number, target: EventTarget): void {
  act(() => {
    target.dispatchEvent(
      new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY, button: 0 })
    );
  });
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
    const { rerender } = renderHook(
      (partEdit: PartEdit | null) =>
        usePagesPointer(
          options({ partEdit, yrsRootStory: partEdit ? `fn:${NOTE_ID}` : 'body', onNoteClick: () => {} })
        ),
      { initialProps: null as PartEdit | null }
    );

    mouse('mousedown', 400, 950, canvasOf());
    expect(selections).toEqual([]);

    act(() => rerender(note));
    expect(selections).toEqual([
      { anchor: NOTE_POSITION, head: NOTE_POSITION, story: `fn:${NOTE_ID}` },
    ]);
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
});
