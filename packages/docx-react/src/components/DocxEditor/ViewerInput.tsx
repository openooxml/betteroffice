import {
  forwardRef,
  memo,
  useCallback,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  type CSSProperties,
} from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { YrsInputRef } from './YrsInput';
import { presentedWorkerFrame } from './internals/layoutProvenance';
import { ViewerSelectionController } from './internals/viewerSelectionController';

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

const MOVE_KEYS = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'];

export interface ViewerInputProps {
  /** Reads the resident document. */
  read: ResidentEngineWorkerClient['documentRead'];
  story: string;
  /** Queries for the presented frame. */
  queries: DisplayListQueries | null;
  document: { isDisplayOnly(): boolean } | null;
  /** The selection or its visibility changed. */
  onSelectionChange(): void;
  onFocusChange?(focused: boolean): void;
}

const ViewerInputComponent = forwardRef<YrsInputRef, ViewerInputProps>(function ViewerInput(
  { read, story, queries, document: documentKey, onSelectionChange, onFocusChange },
  ref
) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const primedRef = useRef<string | null>(null);
  const queriesRef = useRef(queries);
  queriesRef.current = queries;
  const readRef = useRef(read);
  readRef.current = read;
  const onChangeRef = useRef(onSelectionChange);
  onChangeRef.current = onSelectionChange;
  const documentRef = useRef(documentKey);
  const controller = useMemo(() => new ViewerSelectionController({
    get read() { return readRef.current; },
    story,
    queries: () => queriesRef.current,
  }), [story]);

  useLayoutEffect(() => {
    const unsubscribe = controller.subscribe(() => onChangeRef.current());
    return () => {
      unsubscribe();
      controller.reset();
    };
  }, [controller]);
  useLayoutEffect(() => {
    const previous = documentRef.current;
    if (previous === documentKey) return;
    documentRef.current = documentKey;
    if (previous && !(previous.isDisplayOnly() && documentKey && !documentKey.isDisplayOnly())) {
      controller.reset();
    }
  }, [controller, documentKey]);
  useLayoutEffect(() => {
    controller.onFrame(presentedWorkerFrame(queries));
  }, [controller, documentKey, queries]);

  const copy = useCallback((event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    const textarea = textareaRef.current;
    const gesture = controller.currentGesture();
    const ready = controller.settledText(gesture);
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
    const pending = controller.readSelectedText();
    const clipboard = typeof navigator !== 'undefined' ? navigator.clipboard : undefined;
    if (!pending) return;
    event.preventDefault();
    if (!clipboard) {
      void pending.catch(() => {});
      return;
    }
    const text = pending.then((value) => {
      if (!controller.isCurrent(gesture)) throw new Error('Selection gesture changed');
      if (!value) throw new Error('Nothing is selected to copy');
      return value;
    });
    let write: Promise<void>;
    const blob = text.then((value) => new Blob([value], { type: 'text/plain' }));
    void blob.catch(() => {});
    try {
      write = typeof ClipboardItem !== 'undefined' && clipboard.write
        ? clipboard.write([new ClipboardItem({
          'text/plain': blob,
        })])
        : Promise.reject(new Error('No ClipboardItem'));
    } catch (error) {
      write = Promise.reject(error);
    }
    void write.catch(() => text.then((value) => {
      if (!controller.isCurrent(gesture)) throw new Error('Selection gesture changed');
      return clipboard.writeText(value);
    })).catch(() => {});
  }, [controller]);

  const handleKeyDown = useCallback((event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const mod = event.metaKey || event.ctrlKey;
    const key = event.key.toLowerCase();
    if (mod && key === 'a') {
      event.preventDefault();
      controller.selectAll();
    } else if (mod && key === 'c' && !event.shiftKey && !event.altKey) {
      copy(event);
    } else if (!mod && !event.altKey && MOVE_KEYS.includes(event.key) && controller.move(event.key, event.shiftKey)) {
      event.preventDefault();
    }
  }, [controller, copy]);

  const handleCopy = useCallback((event: React.ClipboardEvent<HTMLTextAreaElement>): void => {
    const text = primedRef.current;
    if (text === null) return;
    event.preventDefault();
    event.clipboardData.setData('text/plain', text);
  }, []);

  useImperativeHandle(ref, (): YrsInputRef => ({
    focus: () => textareaRef.current?.focus({ preventScroll: true }),
    blur: () => textareaRef.current?.blur(),
    isFocused: () => typeof document !== 'undefined' && document.activeElement === textareaRef.current,
    flushPendingInput: async () => {},
    runAfterPendingInput: async (operation) => operation(),
    hasPendingInput: () => false,
    beginGesture: () => controller.beginGesture(),
    isGestureCurrent: (gesture) => controller.isCurrent(gesture),
    setSelectionFromDisplay: (anchor, head = anchor, _story, gesture) => controller.select(anchor, head, gesture),
    selectWordAtDisplay: (position) => controller.expand(position, 'word'),
    selectParagraphAtDisplay: (position) => controller.expand(position, 'paragraph'),
    displaySelection: () => controller.displaySelection(),
    keepSelectionInPlace() {},
    applyStoredFormatting() {},
    clearStoredFormatting() {},
    storedFormatting: () => null,
    insertText() {},
    deleteSelection() {},
    selectAll: () => controller.selectAll(),
    readSelectedText: () => controller.readSelectedText(),
  }), [controller]);

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
