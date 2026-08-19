import { useEffect, useRef } from 'react';

const INNER_ESCAPE_LAYER =
  '[role="dialog"], [role="menu"], .oox-hyperlink-popup, [data-docx-escape-layer]';

function hasEditableTarget(event: KeyboardEvent): boolean {
  const target = event.target instanceof Element ? event.target : null;
  const editable = target?.closest('input, textarea, select, [contenteditable]');
  if (!editable || editable.matches('.paged-editor__yrs-input')) return false;
  return editable.getAttribute('contenteditable') !== 'false';
}

function hasInnerEscapeLayer(): boolean {
  return document.querySelector(INNER_ESCAPE_LAYER) != null;
}

/** Runs `onEscape` for an unhandled Escape while `active`. */
export function useEscapeKey(active: boolean, onEscape: () => void): void {
  const onEscapeRef = useRef(onEscape);
  onEscapeRef.current = onEscape;

  useEffect(() => {
    if (!active) return;
    const yieldedEvents = new WeakSet<KeyboardEvent>();
    const markYieldedEvent = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      if (hasEditableTarget(event) || hasInnerEscapeLayer()) yieldedEvents.add(event);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (
        event.key !== 'Escape' ||
        event.isComposing ||
        event.keyCode === 229 ||
        event.defaultPrevented ||
        yieldedEvents.has(event) ||
        hasEditableTarget(event) ||
        hasInnerEscapeLayer()
      ) {
        return;
      }
      event.preventDefault();
      onEscapeRef.current();
    };
    document.addEventListener('keydown', markYieldedEvent, true);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', markYieldedEvent, true);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [active]);
}
