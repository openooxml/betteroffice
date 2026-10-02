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
  DocxTextRange,
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

/** A settled selection's text is read this long after it last changed. */
const SETTLE_MS = 60;

interface ViewerSelection extends DocxDisplayRange {
  /** Bumped on every change; a reply for an older value is dropped. */
  token: number;
  /** The worker version of the layout the positions belong to. */
  version: string;
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
  const tokenRef = useRef(0);
  const queriesRef = useRef(queries);
  queriesRef.current = queries;
  const readRef = useRef(read);
  readRef.current = read;
  const onChangeRef = useRef(onSelectionChange);
  onChangeRef.current = onSelectionChange;
  const textRef = useRef<{ token: number; text: Promise<DocxDisplaySelectionText | null> } | null>(
    null
  );
  const settledRangeRef = useRef<{ range: DocxTextRange; version: string } | null>(null);
  const carryRef = useRef<string | null>(null);
  const reprojectingRef = useRef<string | null>(null);
  const settleTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const primedRef = useRef<string | null>(null);
  const resolvedTextRef = useRef<{ token: number; text: string } | null>(null);

  const presented = (): string | null => presentedWorkerVersion(queriesRef.current);

  const readText = useCallback((selection: ViewerSelection) => {
    const pending = readRef
      .current({
        kind: 'selectionText',
        story,
        anchor: selection.anchor,
        head: selection.head,
        expectVersion: selection.version,
      })
      .then((reply) => (reply.version === selection.version ? reply.value : null))
      .catch(() => null);
    textRef.current = { token: selection.token, text: pending };
    void pending.then((value) => {
      if (!value || selectionRef.current?.token !== selection.token) return;
      resolvedTextRef.current = { token: selection.token, text: value.text };
      if (value.range) settledRangeRef.current = { range: value.range, version: selection.version };
    });
    return pending;
  }, [story]);

  const scheduleSettle = useCallback(() => {
    if (settleTimerRef.current !== null) clearTimeout(settleTimerRef.current);
    settleTimerRef.current = setTimeout(() => {
      settleTimerRef.current = null;
      const selection = selectionRef.current;
      if (selection && selection.anchor !== selection.head && textRef.current?.token !== selection.token) {
        void readText(selection);
      }
    }, SETTLE_MS);
  }, [readText]);

  const setSelection = useCallback(
    (range: DocxDisplayRange, version: string, token = ++tokenRef.current): void => {
      selectionRef.current = { ...range, token, version };
      settledRangeRef.current = null;
      onChangeRef.current();
      scheduleSettle();
    },
    [scheduleSettle]
  );

  const select = useCallback(
    (anchor: number, head: number): number | null => {
      const version = presented();
      if (version === null) return null;
      tokenRef.current += 1;
      setSelection({ anchor, head }, version, tokenRef.current);
      return tokenRef.current;
    },
    [setSelection]
  );

  const readUnit = useCallback(
    (position: number, unit: DocxSelectionUnit, version: string, token: number): void => {
      void readRef
        .current({ kind: 'selectionUnit', story, position, unit, expectVersion: version })
        .then((reply) => {
          const current = selectionRef.current;
          if (!reply.value || reply.version !== version || tokenRef.current !== token) return;
          if (current?.anchor === reply.value.anchor && current.head === reply.value.head) return;
          setSelection(reply.value, version, token);
        })
        .catch(() => {});
    },
    [setSelection, story]
  );

  // The caret lands at once; the presented pages widen it, and the worker settles the unit.
  const expand = useCallback(
    (position: number, unit: 'word' | 'paragraph'): void => {
      const token = select(position, position);
      const version = selectionRef.current?.version;
      if (token === null || !version) return;
      const queries = queriesRef.current;
      if (queries && story === 'body') {
        const local = displayListSelectionUnit(queries, position, unit);
        if (local) setSelection(local, version, token);
      }
      readUnit(position, unit, version, token);
    },
    [readUnit, select, setSelection, story]
  );

  const selectAll = useCallback((): void => {
    const version = presented();
    if (version === null) return;
    readUnit(0, 'story', version, ++tokenRef.current);
  }, [readUnit]);

  const documentRef = useRef(documentKey);
  useEffect(() => {
    const previous = documentRef.current;
    if (previous === documentKey) return;
    documentRef.current = documentKey;
    const selection = selectionRef.current;
    carryRef.current = previous?.isDisplayOnly() && selection ? selection.version : null;
    if (!carryRef.current) {
      selectionRef.current = null;
      settledRangeRef.current = null;
      textRef.current = null;
      onChangeRef.current();
    }
  }, [documentKey]);

  // A frame of another version hides the selection until it maps onto that frame.
  useEffect(() => {
    const selection = selectionRef.current;
    const version = presentedWorkerVersion(queries);
    if (!selection || version === null || selection.version === version) return;
    if (carryRef.current === selection.version) {
      carryRef.current = null;
      selectionRef.current = { ...selection, version };
      textRef.current = null;
      onChangeRef.current();
      scheduleSettle();
      return;
    }
    onChangeRef.current();
    const settled = settledRangeRef.current;
    if (!settled || settled.version !== selection.version || reprojectingRef.current === version) {
      return;
    }
    reprojectingRef.current = version;
    const token = selection.token;
    void readRef
      .current({ kind: 'rangePosition', story, range: settled.range, expectVersion: version })
      .then((reply) => {
        if (!reply.value || reply.version !== version || selectionRef.current?.token !== token) return;
        setSelection(reply.value, version, token);
        settledRangeRef.current = { range: settled.range, version };
      })
      .catch(() => {})
      .finally(() => {
        if (reprojectingRef.current === version) reprojectingRef.current = null;
      });
  }, [queries, scheduleSettle, setSelection, story]);

  useEffect(
    () => () => {
      if (settleTimerRef.current !== null) clearTimeout(settleTimerRef.current);
    },
    []
  );

  const displaySelection = useCallback((): YrsDisplaySelection | null => {
    const selection = selectionRef.current;
    if (!selection || selection.version !== presented()) return null;
    return { anchor: selection.anchor, head: selection.head };
  }, []);

  const selectedText = useCallback((): Promise<string> | null => {
    const selection = selectionRef.current;
    if (!selection || selection.anchor === selection.head || selection.version !== presented()) {
      return null;
    }
    const pending =
      textRef.current?.token === selection.token ? textRef.current.text : readText(selection);
    return pending.then((value) => value?.text ?? '');
  }, [readText]);

  const copy = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
      const textarea = textareaRef.current;
      const pending = selectedText();
      if (!textarea || pending === null) return;
      const resolved =
        resolvedTextRef.current?.token === selectionRef.current?.token
          ? resolvedTextRef.current!.text
          : null;
      if (resolved) {
        primedRef.current = resolved;
        textarea.value = resolved;
        textarea.select();
        setTimeout(() => {
          if (primedRef.current === resolved) primedRef.current = null;
          if (textarea.value === resolved) textarea.value = '';
        });
        return;
      }
      const clipboard = typeof navigator !== 'undefined' ? navigator.clipboard : undefined;
      if (!clipboard) return;
      event.preventDefault();
      const text = Promise.resolve(pending).then((value) => {
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

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
      const mod = event.metaKey || event.ctrlKey;
      const key = event.key.toLowerCase();
      if (mod && key === 'a') {
        event.preventDefault();
        selectAll();
      } else if (mod && key === 'c' && !event.shiftKey && !event.altKey) {
        copy(event);
      }
    },
    [copy, selectAll]
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
