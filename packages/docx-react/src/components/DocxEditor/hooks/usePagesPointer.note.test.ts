/**
 * Clicking into a footnote: the pointer path opens the note it lands in, and
 * the caret that click asked for goes into that note's story once the editor
 * is on it. `partEdit.test.ts` covers the mapping; this shows the wiring.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { createElement, createRef, useState } from 'react';
import type { DisplayListQueries, DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { AlignmentButtons } from '../../ui/AlignmentButtons';
import { MenuDropdown } from '../../ui/MenuDropdown';
import { TableStyleGallery } from '../../ui/TableStyleGallery';
import { FootnotePropertiesDialog } from '../../dialogs/FootnotePropertiesDialog';
import { CommentCard } from '../../sidebar/CommentCard';
import { useEscapeKey } from '../../../hooks/useEscapeKey';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';
import type { YrsInputRef } from '../YrsInput';
import { partEditStory, type PartEdit } from '../partEdit';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
/** y at or past this is the page's footnote area in the fake engine. */
const NOTE_FROM = 900;
const NOTE_A_ID = 1;
const NOTE_ID = 2;
/** what the fake engine resolves anywhere inside the note area */
const NOTE_POSITION = 7;
const UNATTRIBUTED_NOTE_HIT = {
  region: 'footnote',
  noteId: undefined,
  pos: null,
  target: 'none',
} as DisplayListRegionHit;

let host: HTMLDivElement;
let selections: Array<{ anchor: number; head: number; story?: string }>;

function fakeQueries(fixedHit?: DisplayListRegionHit): DisplayListQueries {
  return {
    displayList: { pages: [{ pageIndex: 0, ...PAGE, primitives: [] }] },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number) => {
      if (fixedHit) return fixedHit;
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

function dispatchEscape(target: EventTarget, init: KeyboardEventInit = {}): KeyboardEvent {
  const event = new KeyboardEvent('keydown', {
    key: 'Escape',
    bubbles: true,
    cancelable: true,
    ...init,
  });
  act(() => target.dispatchEvent(event));
  return event;
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

  test('a click in the body leaves the note and places the body caret there', () => {
    let left = 0;
    const options = stableOptions();
    renderHook(() => {
      const [partEdit, setPartEdit] = useState<PartEdit | null>(note);
      return usePagesPointer(
        options({
          partEdit,
          yrsRootStory: partEditStory(partEdit),
          onBodyClick: () => {
            left += 1;
            setPartEdit(null);
          },
        })
      );
    });

    mouse('mousedown', 400, 500, canvasOf());

    expect(left).toBe(1);
    expect(selections).toEqual([{ anchor: 1, head: 1, story: 'body' }]);
  });

  test('a click in the body leaves a band and places the body caret there', () => {
    let left = 0;
    const options = stableOptions();
    renderHook(() => {
      const [partEdit, setPartEdit] = useState<PartEdit | null>({
        kind: 'header',
        rId: 'rId7',
      });
      return usePagesPointer(
        options({
          partEdit,
          yrsRootStory: partEditStory(partEdit),
          onBodyClick: () => {
            left += 1;
            setPartEdit(null);
          },
        })
      );
    });

    mouse('mousedown', 400, 500, canvasOf());

    expect(left).toBe(1);
    expect(selections).toEqual([{ anchor: 1, head: 1, story: 'body' }]);
  });

  test('a body click can leave a note whose story vanished', () => {
    let left = 0;
    const options = stableOptions();
    const projectionFor = options().getYrsPositionProjection;
    renderHook(() => {
      const [partEdit, setPartEdit] = useState<PartEdit | null>(note);
      return usePagesPointer(
        options({
          partEdit,
          yrsRootStory: partEditStory(partEdit),
          getYrsPositionProjection: (story) =>
            story === `fn:${NOTE_ID}` ? null : projectionFor(story),
          onBodyClick: () => {
            left += 1;
            setPartEdit(null);
          },
        })
      );
    });

    mouse('mousedown', 400, 500, canvasOf());

    expect(left).toBe(1);
    expect(selections).toEqual([{ anchor: 1, head: 1, story: 'body' }]);
  });

  test('an unattributed note-area hit is inert while the body is active', () => {
    const options = stableOptions();
    renderHook(() =>
      usePagesPointer(
        options({ displayListQueries: fakeQueries(UNATTRIBUTED_NOTE_HIT), onNoteClick: () => {} })
      )
    );

    mouse('mousedown', 400, 950, canvasOf());

    expect(selections).toEqual([]);
  });

  test('an unattributed note-area hit does not close an open note', () => {
    let left = 0;
    const options = stableOptions();
    renderHook(() =>
      usePagesPointer(
        options({
          partEdit: note,
          yrsRootStory: `fn:${NOTE_ID}`,
          displayListQueries: fakeQueries(UNATTRIBUTED_NOTE_HIT),
          onBodyClick: () => (left += 1),
        })
      )
    );

    mouse('mousedown', 400, 950, canvasOf());

    expect(left).toBe(0);
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

describe('closing an open part with Escape', () => {
  test('ignores composition Escape events', () => {
    let closed = 0;
    renderHook(() => useEscapeKey(true, () => (closed += 1)));

    dispatchEscape(document, { isComposing: true });
    const legacy = new KeyboardEvent('keydown', {
      key: 'Escape',
      bubbles: true,
      cancelable: true,
    });
    Object.defineProperty(legacy, 'keyCode', { value: 229 });
    act(() => document.dispatchEvent(legacy));

    expect(closed).toBe(0);
  });

  test('lets an open menu consume Escape without closing the part', () => {
    let closed = 0;
    const view = render(createElement(EscapeMenuHarness, { onEscape: () => (closed += 1) }));
    const trigger = view.getByRole('button', { name: 'File' });
    act(() => trigger.click());

    expect(view.getByText('Open')).toBeTruthy();
    dispatchEscape(trigger);

    expect(view.queryByText('Open')).toBeNull();
    expect(closed).toBe(0);
  });

  test('lets an open alignment dropdown consume Escape without closing the part', () => {
    let closed = 0;
    const view = render(
      createElement(EscapeAlignmentHarness, { onEscape: () => (closed += 1) })
    );
    const trigger = view.getByTestId('toolbar-alignment');
    act(() => trigger.click());

    expect(view.getByTestId('alignment-left')).toBeTruthy();
    const input = document.createElement('textarea');
    input.className = 'paged-editor__yrs-input';
    host.append(input);
    dispatchEscape(input);

    expect(view.queryByTestId('alignment-left')).toBeNull();
    expect(closed).toBe(0);
    input.remove();
  });

  test('lets the table style gallery consume Escape without closing the part', () => {
    let closed = 0;
    const view = render(
      createElement(EscapeTableStyleHarness, { onEscape: () => (closed += 1) })
    );
    const trigger = view.container.querySelector('button');
    expect(trigger).not.toBeNull();
    act(() => trigger!.click());

    expect(view.container.querySelector('[data-docx-escape-layer]')).not.toBeNull();
    dispatchEscape(trigger!);

    expect(view.container.querySelector('[data-docx-escape-layer]')).toBeNull();
    expect(closed).toBe(0);
  });

  test('lets the footnote properties dialog consume Escape without closing the part', () => {
    let closed = 0;
    let dialogClosed = 0;
    const view = render(
      createElement(EscapeFootnotePropertiesHarness, {
        onEscape: () => (closed += 1),
        onClose: () => (dialogClosed += 1),
      })
    );

    dispatchEscape(view.getByRole('dialog'));

    expect(view.queryByRole('dialog')).toBeNull();
    expect(dialogClosed).toBe(1);
    expect(closed).toBe(0);
  });

  test('lets a comment menu consume Escape without closing the part', () => {
    let closed = 0;
    const view = render(
      createElement(EscapeCommentMenuHarness, { onEscape: () => (closed += 1) })
    );
    const trigger = view.container.querySelector<HTMLButtonElement>('[aria-haspopup="menu"]');
    expect(trigger).not.toBeNull();
    act(() => trigger!.click());

    expect(view.getByRole('menu')).toBeTruthy();
    dispatchEscape(trigger!);

    expect(view.queryByRole('menu')).toBeNull();
    expect(closed).toBe(0);
  });

  test('plain Escape from the editor input closes the part', () => {
    let closed = 0;
    const input = document.createElement('textarea');
    input.className = 'paged-editor__yrs-input';
    host.append(input);
    renderHook(() => useEscapeKey(true, () => (closed += 1)));

    const event = dispatchEscape(input);

    expect(closed).toBe(1);
    expect(event.defaultPrevented).toBe(true);
    input.remove();
  });

  test('ignores Escape from another input', () => {
    let closed = 0;
    const input = document.createElement('input');
    host.append(input);
    renderHook(() => useEscapeKey(true, () => (closed += 1)));

    dispatchEscape(input);

    expect(closed).toBe(0);
    input.remove();
  });

  test('ignores an already prevented Escape', () => {
    let closed = 0;
    renderHook(() => useEscapeKey(true, () => (closed += 1)));
    const event = new KeyboardEvent('keydown', {
      key: 'Escape',
      bubbles: true,
      cancelable: true,
    });
    event.preventDefault();

    act(() => document.dispatchEvent(event));

    expect(closed).toBe(0);
  });

  test('closes an active part without mounted chrome', () => {
    let closed = 0;
    renderHook(() => useEscapeKey(true, () => (closed += 1)));

    dispatchEscape(document);

    expect(closed).toBe(1);
  });
});

function EscapeMenuHarness({ onEscape }: { onEscape: () => void }) {
  useEscapeKey(true, onEscape);
  return createElement(MenuDropdown, {
    label: 'File',
    items: [{ label: 'Open', onClick: () => {} }],
  });
}

function EscapeAlignmentHarness({ onEscape }: { onEscape: () => void }) {
  useEscapeKey(true, onEscape);
  return createElement(AlignmentButtons);
}

function EscapeTableStyleHarness({ onEscape }: { onEscape: () => void }) {
  useEscapeKey(true, onEscape);
  return createElement(TableStyleGallery, { onAction: () => {} });
}

function EscapeFootnotePropertiesHarness({
  onEscape,
  onClose,
}: {
  onEscape: () => void;
  onClose: () => void;
}) {
  useEscapeKey(true, onEscape);
  const [open, setOpen] = useState(true);
  return createElement(FootnotePropertiesDialog, {
    isOpen: open,
    onClose: () => {
      setOpen(false);
      onClose();
    },
    onApply: () => {},
  });
}

function EscapeCommentMenuHarness({ onEscape }: { onEscape: () => void }) {
  useEscapeKey(true, onEscape);
  return createElement(CommentCard, {
    comment: { id: 1, author: 'Reviewer', content: [] },
    replies: [],
    isExpanded: true,
    onToggleExpand: () => {},
    measureRef: () => {},
  });
}
