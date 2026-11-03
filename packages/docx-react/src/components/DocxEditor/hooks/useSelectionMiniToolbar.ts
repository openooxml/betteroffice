import { useCallback, useEffect, useRef, useState } from 'react';
import type { SelectionToolbarAnchorRect } from '../../SelectionMiniToolbar';
import type { DocxSelectionToolbarItem } from '../../DocxEditor';

/** Items shown when `selectionToolbar` is `true` or names no items. */
export const DEFAULT_SELECTION_TOOLBAR_ITEMS: readonly DocxSelectionToolbarItem[] = [
  'clearFormatting',
  'pastePlainText',
];

/** The items to show for a `selectionToolbar` prop value, or null when it is off. */
export function resolveSelectionToolbarItems(
  selectionToolbar: boolean | { items?: readonly DocxSelectionToolbarItem[] } | undefined
): readonly DocxSelectionToolbarItem[] | null {
  if (selectionToolbar === false) return null;
  if (selectionToolbar === true || selectionToolbar === undefined) return DEFAULT_SELECTION_TOOLBAR_ITEMS;
  return selectionToolbar.items ?? DEFAULT_SELECTION_TOOLBAR_ITEMS;
}

/** Quiet time after the last selection change or pointer release before the toolbar shows. */
export const SELECTION_TOOLBAR_SETTLE_MS = 350;

export interface UseSelectionMiniToolbarOptions {
  enabled: boolean;
  /** Editor root element; the toolbar renders inside it. */
  rootRef: React.RefObject<HTMLElement | null>;
  /** The editor's scroll container, when it is not the window. */
  scrollRef?: React.RefObject<HTMLElement | null>;
  /** Identity of the current non-empty selection, or null when collapsed or absent. */
  selectionKey: string | null;
  /** Client-space union of the selection rects, read when the toolbar is about to show. */
  getAnchorRect: () => SelectionToolbarAnchorRect | null;
  /** Changes whenever the document changes; a change hides the toolbar. */
  documentKey?: unknown;
  /** Whether an event target outside `rootRef` still belongs to the editor (e.g. its portaled text input). */
  isEditorTarget?: (target: EventTarget | null) => boolean;
  settleMs?: number;
}

export interface UseSelectionMiniToolbarReturn {
  visible: boolean;
  anchorRect: SelectionToolbarAnchorRect | null;
  toolbarRef: (element: HTMLDivElement | null) => void;
  dismiss: (returnFocus: boolean) => void;
  focusToolbar: () => void;
}

function isTypingKey(event: KeyboardEvent): boolean {
  if (event.altKey && event.key === 'F10') return false;
  if (event.ctrlKey || event.metaKey) return event.key.toLowerCase() === 'v' || event.key.toLowerCase() === 'x';
  if (event.altKey) return false;
  return event.key.length === 1 || event.key === 'Enter' || event.key === 'Backspace' || event.key === 'Delete';
}

/**
 * Decides when the selection mini toolbar is shown: only once a selection
 * has settled (after pointer up, never between the clicks of a double or
 * triple click, and after keyboard selection stops changing), and hidden
 * by typing, scrolling, resizing the window, Escape, focus loss, a collapsed selection or a
 * document change.
 */
export function useSelectionMiniToolbar({
  enabled,
  rootRef,
  scrollRef,
  selectionKey,
  getAnchorRect,
  documentKey,
  isEditorTarget,
  settleMs = SELECTION_TOOLBAR_SETTLE_MS,
}: UseSelectionMiniToolbarOptions): UseSelectionMiniToolbarReturn {
  const [visible, setVisible] = useState(false);
  const [anchorRect, setAnchorRect] = useState<SelectionToolbarAnchorRect | null>(null);
  const toolbarElement = useRef<HTMLDivElement | null>(null);
  const pointerDown = useRef(false);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const suppressedKey = useRef<string | null>(null);
  const selectionKeyRef = useRef(selectionKey);
  selectionKeyRef.current = selectionKey;
  const isEditorTargetRef = useRef(isEditorTarget);
  isEditorTargetRef.current = isEditorTarget;
  const getAnchorRectRef = useRef(getAnchorRect);
  getAnchorRectRef.current = getAnchorRect;
  const enabledRef = useRef(enabled);
  enabledRef.current = enabled;

  const clearTimer = useCallback(() => {
    if (timer.current !== null) clearTimeout(timer.current);
    timer.current = null;
  }, []);

  const hide = useCallback(() => {
    clearTimer();
    setVisible(false);
  }, [clearTimer]);

  const dismiss = useCallback(
    (_returnFocus: boolean) => {
      suppressedKey.current = selectionKeyRef.current;
      hide();
    },
    [hide]
  );

  const scheduleShow = useCallback(() => {
    clearTimer();
    if (!enabledRef.current || pointerDown.current) return;
    const key = selectionKeyRef.current;
    if (key === null || suppressedKey.current === key) return;
    timer.current = setTimeout(() => {
      timer.current = null;
      if (!enabledRef.current || pointerDown.current) return;
      if (selectionKeyRef.current === null || suppressedKey.current === selectionKeyRef.current) return;
      const rect = getAnchorRectRef.current();
      if (!rect) return;
      setAnchorRect(rect);
      setVisible(true);
    }, settleMs);
  }, [clearTimer, settleMs]);

  const toolbarRef = useCallback((element: HTMLDivElement | null) => {
    toolbarElement.current = element;
  }, []);

  const focusToolbar = useCallback(() => {
    toolbarElement.current?.querySelector<HTMLElement>('button')?.focus();
  }, []);

  // A collapsed or removed selection always hides the toolbar; a changed
  // selection must settle again before it shows.
  useEffect(() => {
    if (!enabled || selectionKey === null) {
      suppressedKey.current = null;
      hide();
      return;
    }
    if (suppressedKey.current !== null && suppressedKey.current !== selectionKey) {
      suppressedKey.current = null;
    }
    setVisible(false);
    scheduleShow();
  }, [enabled, selectionKey, hide, scheduleShow]);

  // A document change hides the toolbar until the selection changes again.
  const previousDocumentKey = useRef(documentKey);
  useEffect(() => {
    if (Object.is(previousDocumentKey.current, documentKey)) return;
    previousDocumentKey.current = documentKey;
    suppressedKey.current = selectionKeyRef.current;
    hide();
  }, [documentKey, hide]);

  useEffect(() => {
    if (!enabled || typeof document === 'undefined') return;
    const inToolbar = (target: EventTarget | null) =>
      target instanceof Node && !!toolbarElement.current?.contains(target);
    const inRoot = (target: EventTarget | null) =>
      (target instanceof Node && !!rootRef.current?.contains(target)) || !!isEditorTargetRef.current?.(target);

    const onPointerDown = (event: Event) => {
      if (inToolbar(event.target) || !inRoot(event.target)) return;
      pointerDown.current = true;
      hide();
    };
    const onPointerUp = () => {
      if (!pointerDown.current) return;
      pointerDown.current = false;
      scheduleShow();
    };
    const onPointerCancel = () => {
      pointerDown.current = false;
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (inToolbar(event.target)) return;
      if (event.altKey && event.key === 'F10') {
        if (visible && inRoot(event.target)) {
          event.preventDefault();
          event.stopPropagation();
          focusToolbar();
        }
        return;
      }
      if (!inRoot(event.target)) return;
      if (event.key === 'Escape') {
        if (visible) dismiss(false);
        return;
      }
      if (isTypingKey(event)) {
        suppressedKey.current = selectionKeyRef.current;
        hide();
      }
    };
    const onScroll = (event: Event) => {
      if (inToolbar(event.target)) return;
      const scroller = scrollRef?.current;
      // Window scrolls arrive with the document as target.
      if (scroller && event.target !== scroller && event.target !== document) return;
      suppressedKey.current = selectionKeyRef.current;
      hide();
    };
    const onResize = () => {
      if (!visible) return;
      suppressedKey.current = selectionKeyRef.current;
      hide();
    };
    const onFocusOut = (event: FocusEvent) => {
      if (!inRoot(event.target) && !inToolbar(event.target)) return;
      const next = event.relatedTarget as Node | null;
      if (next && (inRoot(next) || inToolbar(next))) return;
      hide();
    };

    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('mousedown', onPointerDown, true);
    document.addEventListener('pointerup', onPointerUp, true);
    document.addEventListener('mouseup', onPointerUp, true);
    document.addEventListener('pointercancel', onPointerCancel, true);
    document.addEventListener('keydown', onKeyDown, true);
    document.addEventListener('focusout', onFocusOut);
    document.addEventListener('scroll', onScroll, { capture: true, passive: true });
    window.addEventListener('scroll', onScroll, { passive: true });
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('mousedown', onPointerDown, true);
      document.removeEventListener('pointerup', onPointerUp, true);
      document.removeEventListener('mouseup', onPointerUp, true);
      document.removeEventListener('pointercancel', onPointerCancel, true);
      document.removeEventListener('keydown', onKeyDown, true);
      document.removeEventListener('focusout', onFocusOut);
      document.removeEventListener('scroll', onScroll, { capture: true });
      window.removeEventListener('scroll', onScroll);
      window.removeEventListener('resize', onResize);
    };
  }, [enabled, rootRef, scrollRef, visible, hide, scheduleShow, dismiss, focusToolbar]);

  return { visible: enabled && visible, anchorRect, toolbarRef, dismiss, focusToolbar };
}
