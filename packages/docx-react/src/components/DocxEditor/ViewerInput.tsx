import {
  forwardRef,
  memo,
  useCallback,
  useEffect,
  useImperativeHandle,
  useRef,
  type CSSProperties,
} from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type {
  DocxDisplayRange,
  DocxDisplaySelectionText,
  DocxSelectionUnit,
  ResidentEngineWorkerClient,
} from '@betteroffice/docx/yrs';
import type { YrsDisplaySelection, YrsInputRef } from './YrsInput';
import { displayListSelectionUnit } from './internals/viewerSelectionUnits';
import { presentedWorkerVersion } from './internals/layoutProvenance';

const STYLE: CSSProperties = {
  position: 'fixed',
  left: 0,
  top: 0,
  width: '1px',
  height: '1px',
  padding: 0,
  margin: 0,
  border: 0,
  outline: 0,
  opacity: 0,
  overflow: 'hidden',
  resize: 'none',
  zIndex: -1,
  background: 'transparent',
  color: 'transparent',
  caretColor: 'transparent',
};

/** A copy waits up to TEXT_RETRIES × RETRY_MS for its selection to land on the presented frame. */
const RETRY_MS = 60;
const TEXT_RETRIES = 100;

const MOVE_KEYS = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'];

interface ViewerSelection extends DocxDisplayRange {
  /** The gesture that made it; a unit reply for an older gesture is dropped. */
  gesture: number;
  /** Bumped on every change; text read for an older revision is not this selection's. */
  revision: number;
  /** The worker version of the layout the positions belong to. */
  version: string;
  /** Made on a display-only preview, whose first pages the document lays out the same. */
  preview: boolean;
}

interface Capture {
  revision: number;
  read: Promise<DocxDisplaySelectionText | null>;
  value?: DocxDisplaySelectionText | null;
}

export interface ViewerInputProps {
  /** Reads the document the resident worker holds. */
  read: ResidentEngineWorkerClient['documentRead'];
  /** The root story the pages show. */
  story: string;
  /** The display list queries of the presented frame. */
  queries: DisplayListQueries | null;
  /**
   * The open document. A new one drops the selection, except after a display-only preview,
   * whose first pages the document lays out the same.
   */
  document: { isDisplayOnly(): boolean } | null;
  /** The selection or its visibility changed. */
  onSelectionChange(): void;
  onFocusChange?(focused: boolean): void;
}

/**
 * The input surface of a viewer session: it holds the selection in display positions of the
 * presented frame and reads text from the document worker, never from a document on this thread.
 */
const ViewerInputComponent = forwardRef<YrsInputRef, ViewerInputProps>(function ViewerInput(
  { read, story, queries, document: documentKey, onSelectionChange, onFocusChange },
  ref
) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const selectionRef = useRef<ViewerSelection | null>(null);
  const gestureRef = useRef(0);
  const revisionRef = useRef(0);
  const captureRef = useRef<Capture | null>(null);
  const pendingRef = useRef<Promise<void> | null>(null);
  const capturingRef = useRef(false);
  const captureLatestRef = useRef<() => void>(() => {});
  const goalXRef = useRef<number | undefined>(undefined);
  const primedRef = useRef<string | null>(null);
  const queriesRef = useRef(queries);
  queriesRef.current = queries;
  const readRef = useRef(read);
  readRef.current = read;
  const onChangeRef = useRef(onSelectionChange);
  onChangeRef.current = onSelectionChange;
  const documentRef = useRef(documentKey);
  const reprojectRef = useRef<() => void>(() => {});

  const presented = (): string | null => presentedWorkerVersion(queriesRef.current);

  const capture = useCallback(
    (selection: ViewerSelection): Capture => {
      const current = captureRef.current;
      if (current?.revision === selection.revision) return current;
      const next: Capture = {
        revision: selection.revision,
        read: readRef
          .current({
            kind: 'selectionText',
            story,
            anchor: selection.anchor,
            head: selection.head,
            expectVersion: selection.version,
          })
          .then((reply) => (reply.version === selection.version ? reply.value : null))
          .catch(() => null),
      };
      captureRef.current = next;
      void next.read.then((value) => {
        next.value = value;
        // The document may have moved on while the text was read.
        if (selectionRef.current?.revision === selection.revision) reprojectRef.current();
      });
      return next;
    },
    [story]
  );

  // One capture at a time; the one landing reads the selection made meanwhile.
  const captureLatest = useCallback(() => {
    const selection = selectionRef.current;
    if (capturingRef.current || !selection || selection.version !== presented()) return;
    capturingRef.current = true;
    void capture(selection).read.finally(() => {
      capturingRef.current = false;
      const latest = selectionRef.current;
      if (!latest || latest.revision === selection.revision) return;
      if (latest.version === presented()) {
        captureLatestRef.current();
      } else if (!latest.preview) {
        // Made on a version the document has since left, before it could be read: it cannot follow.
        selectionRef.current = null;
        onChangeRef.current();
      }
    });
  }, [capture]);
  captureLatestRef.current = captureLatest;

  const setSelection = useCallback(
    (range: DocxDisplayRange, version: string, gesture: number): void => {
      selectionRef.current = {
        anchor: range.anchor,
        head: range.head,
        gesture,
        revision: ++revisionRef.current,
        version,
        preview: documentRef.current?.isDisplayOnly() ?? false,
      };
      onChangeRef.current();
      captureLatest();
    },
    [captureLatest]
  );

  const select = useCallback(
    (anchor: number, head: number): number | null => {
      const version = presented();
      if (version === null) return null;
      const gesture = ++gestureRef.current;
      goalXRef.current = undefined;
      setSelection({ anchor, head }, version, gesture);
      return gesture;
    },
    [setSelection]
  );

  const readUnit = useCallback(
    (position: number, unit: DocxSelectionUnit, version: string, gesture: number): Promise<void> =>
      readRef
        .current({ kind: 'selectionUnit', story, position, unit, expectVersion: version })
        .then((reply) => {
          const current = selectionRef.current;
          if (!reply.value || reply.version !== version || gestureRef.current !== gesture) return;
          if (current?.anchor === reply.value.anchor && current.head === reply.value.head) return;
          setSelection(reply.value, version, gesture);
        })
        .catch(() => {}),
    [setSelection, story]
  );

  const track = useCallback((pending: Promise<void>): void => {
    pendingRef.current = pending;
    void pending.finally(() => {
      if (pendingRef.current === pending) pendingRef.current = null;
    });
  }, []);

  // The caret lands at once; the presented pages widen it, and the worker settles the unit.
  const expand = useCallback(
    (position: number, unit: 'word' | 'paragraph'): void => {
      const gesture = select(position, position);
      const version = selectionRef.current?.version;
      if (gesture === null || !version) return;
      const shown = queriesRef.current;
      if (shown && story === 'body') {
        const local = displayListSelectionUnit(shown, position, unit);
        if (local) setSelection(local, version, gesture);
      }
      track(readUnit(position, unit, version, gesture));
    },
    [readUnit, select, setSelection, story, track]
  );

  const selectAll = useCallback((): void => {
    const version = presented();
    if (version === null) return;
    goalXRef.current = undefined;
    track(readUnit(0, 'story', version, ++gestureRef.current));
  }, [readUnit, track]);

  const drop = (): void => {
    selectionRef.current = null;
    onChangeRef.current();
  };

  // A selection from another version maps onto the presented frame through its sticky ends.
  const reproject = useCallback((): void => {
    const selection = selectionRef.current;
    const version = presented();
    if (!selection || version === null || selection.version === version) return;
    if (selection.preview) {
      setSelection(selection, version, selection.gesture);
      return;
    }
    const captured = captureRef.current;
    if (captured?.revision !== selection.revision || captured.value === undefined) return;
    const sticky = captured.value?.sticky;
    if (!sticky) {
      drop();
      return;
    }
    const revision = selection.revision;
    void readRef
      .current({
        kind: 'stickyPosition',
        story,
        anchor: sticky.anchor,
        head: sticky.head,
        expectVersion: version,
      })
      .then((reply) => {
        if (selectionRef.current?.revision !== revision || reply.version !== presented()) return;
        if (reply.value) setSelection(reply.value, reply.version, selection.gesture);
        else drop();
      })
      .catch(() => {});
  }, [setSelection, story]);
  reprojectRef.current = reproject;

  useEffect(() => {
    const selection = selectionRef.current;
    if (!selection || selection.version === presentedWorkerVersion(queries)) return;
    // Hidden until it maps onto this frame; one never captured cannot follow the change.
    onChangeRef.current();
    if (
      !selection.preview &&
      !capturingRef.current &&
      captureRef.current?.revision !== selection.revision
    ) {
      selectionRef.current = null;
      return;
    }
    reproject();
  }, [queries, reproject]);

  useEffect(() => {
    const previous = documentRef.current;
    if (previous === documentKey || documentKey === null) return;
    documentRef.current = documentKey;
    if (previous && !previous.isDisplayOnly()) {
      selectionRef.current = null;
      captureRef.current = null;
      onChangeRef.current();
    }
  }, [documentKey]);

  const displaySelection = useCallback((): YrsDisplaySelection | null => {
    const selection = selectionRef.current;
    if (!selection || selection.version !== presented()) return null;
    return { anchor: selection.anchor, head: selection.head };
  }, []);

  /** The selection's text, once a select-all or unit read still on its way has landed. */
  const selectedText = useCallback((): Promise<string> | null => {
    // A selection on the preview, or one a new version maps on, reads once it lands on the frame.
    const textOf = async (attempt = 0): Promise<string> => {
      const selection = selectionRef.current;
      if (!selection || selection.anchor === selection.head) return '';
      if (selection.version === presented()) {
        const value = await capture(selection).read;
        if (value) return value.text;
      }
      const settled =
        !selection.preview && selection.version === presented() && selectionRef.current === selection;
      if (attempt >= TEXT_RETRIES || settled) return '';
      await new Promise((resolve) => setTimeout(resolve, RETRY_MS));
      return textOf(attempt + 1);
    };
    const pending = pendingRef.current;
    if (pending) return pending.then(textOf);
    const selection = selectionRef.current;
    return selection && selection.anchor !== selection.head ? textOf() : null;
  }, [capture]);

  const copy = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
      const textarea = textareaRef.current;
      const selection = selectionRef.current;
      const captured = captureRef.current;
      const ready =
        !pendingRef.current &&
        selection &&
        selection.version === presented() &&
        captured?.revision === selection.revision
          ? captured.value?.text
          : undefined;
      if (textarea && ready) {
        primedRef.current = ready;
        textarea.value = ready;
        textarea.select();
        setTimeout(() => {
          if (primedRef.current === ready) primedRef.current = null;
          if (textarea.value === ready) textarea.value = '';
        });
        return;
      }
      const pending = selectedText();
      const clipboard = typeof navigator !== 'undefined' ? navigator.clipboard : undefined;
      if (!pending || !clipboard) return;
      event.preventDefault();
      const text = pending.then((value) => {
        if (!value) throw new Error('Nothing is selected to copy');
        return value;
      });
      const write =
        typeof ClipboardItem !== 'undefined' && clipboard.write
          ? clipboard.write([
              new ClipboardItem({
                'text/plain': text.then((value) => new Blob([value], { type: 'text/plain' })),
              }),
            ])
          : Promise.reject(new Error('No ClipboardItem'));
      void write.catch(() => text.then((value) => clipboard.writeText(value))).catch(() => {});
    },
    [selectedText]
  );

  /** Moves the head over the presented frame's lines; false when there is nothing to move. */
  const move = useCallback(
    (key: string, extend: boolean): boolean => {
      const selection = selectionRef.current;
      const shown = queriesRef.current;
      if (!selection || !shown?.isReady() || selection.version !== presented()) return false;
      const from = Math.min(selection.anchor, selection.head);
      const to = Math.max(selection.anchor, selection.head);
      let head = selection.head;
      let goalX: number | undefined;
      if ((key === 'ArrowLeft' || key === 'ArrowRight') && !extend && from !== to) {
        head = key === 'ArrowLeft' ? from : to;
      } else if (key === 'ArrowLeft' || key === 'ArrowRight') {
        const step = key === 'ArrowLeft' ? -1 : 1;
        for (let next = head + step, tries = 0; next >= 0 && tries < 4; next += step, tries += 1) {
          if (shown.visualLineAtPosition(next)) {
            head = next;
            break;
          }
        }
      } else if (key === 'ArrowUp' || key === 'ArrowDown') {
        const moved = shown.verticalMove(head, key === 'ArrowUp' ? 'up' : 'down', goalXRef.current);
        if (!moved) return true;
        head = moved.position;
        goalX = moved.goalX;
      } else {
        const line = shown.visualLineAtPosition(head);
        if (!line) return true;
        head = key === 'Home' ? line.from : line.to;
      }
      setSelection({ anchor: extend ? selection.anchor : head, head }, selection.version, ++gestureRef.current);
      goalXRef.current = goalX;
      return true;
    },
    [setSelection]
  );

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
      const mod = event.metaKey || event.ctrlKey;
      const key = event.key.toLowerCase();
      if (mod && key === 'a') {
        event.preventDefault();
        selectAll();
      } else if (mod && key === 'c' && !event.shiftKey && !event.altKey) {
        copy(event);
      } else if (!mod && !event.altKey && MOVE_KEYS.includes(event.key) && move(event.key, event.shiftKey)) {
        event.preventDefault();
      }
    },
    [copy, move, selectAll]
  );

  const handleCopy = useCallback((event: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const text = primedRef.current;
    primedRef.current = null;
    if (!text) return;
    event.preventDefault();
    event.clipboardData.setData('text/plain', text);
  }, []);

  useImperativeHandle(
    ref,
    (): YrsInputRef => ({
      focus: () => textareaRef.current?.focus({ preventScroll: true }),
      blur: () => textareaRef.current?.blur(),
      isFocused: () =>
        typeof document !== 'undefined' && document.activeElement === textareaRef.current,
      flushPendingInput: async () => {},
      runAfterPendingInput: async (operation) => operation(),
      hasPendingInput: () => false,
      setSelectionFromDisplay(anchor, head = anchor) {
        select(anchor, head);
      },
      selectWordAtDisplay(position) {
        expand(position, 'word');
      },
      selectParagraphAtDisplay(position) {
        expand(position, 'paragraph');
      },
      displaySelection,
      keepSelectionInPlace() {},
      applyStoredFormatting() {},
      clearStoredFormatting() {},
      storedFormatting: () => null,
      insertText() {},
      deleteSelection() {},
      selectAll,
      readSelectedText: selectedText,
    }),
    [displaySelection, expand, select, selectAll, selectedText]
  );

  const textarea = (
    <textarea
      ref={textareaRef}
      className="paged-editor__yrs-input paged-editor__hidden-pm ProseMirror"
      data-testid="yrs-input"
      data-yrs-story={story}
      aria-label="Document input"
      readOnly
      rows={1}
      style={STYLE}
      onKeyDown={handleKeyDown}
      onCopy={handleCopy}
      onFocus={(event) => {
        event.currentTarget.classList.add('ProseMirror-focused');
        onFocusChange?.(true);
      }}
      onBlur={(event) => {
        event.currentTarget.classList.remove('ProseMirror-focused');
        onFocusChange?.(false);
      }}
    />
  );
  return typeof document !== 'undefined' && document.body
    ? createPortal(textarea, document.body)
    : textarea;
});

export const ViewerInput = memo(ViewerInputComponent);
