import { useCallback, useRef } from 'react';
import type { KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';

export interface RovingFocusOptions {
  /** Element whose matching descendants take part, in DOM order. */
  container: RefObject<HTMLElement | null>;
  /** Selector for focusable items; nested menus are excluded by `scope`. */
  itemSelector: string;
  /** Items belong to the nearest ancestor matching this selector. */
  scope?: string;
  orientation?: 'vertical' | 'horizontal';
  /** Label used for typeahead matching; defaults to the text content. */
  labelOf?(item: HTMLElement): string;
}

export interface RovingFocus {
  items(): HTMLElement[];
  focusFirst(): void;
  focusLast(): void;
  /** Handles arrows, Home/End and typeahead; returns whether it moved focus. */
  onKeyDown(event: ReactKeyboardEvent | KeyboardEvent): boolean;
}

const TYPEAHEAD_RESET_MS = 500;

/** Keyboard focus movement between the items of one menu or toolbar. */
export function useRovingFocus(options: RovingFocusOptions): RovingFocus {
  const optionsRef = useRef(options);
  optionsRef.current = options;
  const typeahead = useRef({ buffer: '', at: 0 });

  const items = useCallback((): HTMLElement[] => {
    const { container, itemSelector, scope } = optionsRef.current;
    const root = container.current;
    if (!root) return [];
    return Array.from(root.querySelectorAll<HTMLElement>(itemSelector)).filter(
      (item) => !scope || item.closest(scope) === root
    );
  }, []);

  const focusAt = useCallback(
    (index: number) => {
      const list = items();
      if (list.length === 0) return;
      list[(index + list.length) % list.length].focus();
    },
    [items]
  );

  const onKeyDown = useCallback(
    (event: ReactKeyboardEvent | KeyboardEvent): boolean => {
      const list = items();
      if (list.length === 0) return false;
      const current = list.findIndex((item) => item === document.activeElement);
      const vertical = (optionsRef.current.orientation ?? 'vertical') === 'vertical';
      const next = vertical ? 'ArrowDown' : 'ArrowRight';
      const previous = vertical ? 'ArrowUp' : 'ArrowLeft';
      let target: number | null = null;
      if (event.key === next) target = current + 1;
      else if (event.key === previous) target = current < 0 ? list.length - 1 : current - 1;
      else if (event.key === 'Home') target = 0;
      else if (event.key === 'End') target = list.length - 1;
      else if (
        event.key.length === 1 &&
        event.key !== ' ' &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.altKey
      ) {
        const now = Date.now();
        const state = typeahead.current;
        state.buffer =
          now - state.at > TYPEAHEAD_RESET_MS ? event.key.toLowerCase() : state.buffer + event.key.toLowerCase();
        state.at = now;
        const labelOf =
          optionsRef.current.labelOf ?? ((item: HTMLElement) => item.textContent ?? '');
        const ordered = [...list.slice(current + 1), ...list.slice(0, current + 1)];
        const repeated = [...state.buffer].every((char) => char === state.buffer[0]);
        const match =
          ordered.find((item) => labelOf(item).trim().toLowerCase().startsWith(state.buffer)) ??
          (repeated
            ? ordered.find((item) => labelOf(item).trim().toLowerCase().startsWith(state.buffer[0]))
            : undefined);
        if (!match) return false;
        event.preventDefault();
        match.focus();
        return true;
      }
      if (target === null) return false;
      event.preventDefault();
      focusAt(target);
      return true;
    },
    [focusAt, items]
  );

  return {
    items,
    focusFirst: useCallback(() => focusAt(0), [focusAt]),
    focusLast: useCallback(() => focusAt(-1), [focusAt]),
    onKeyDown,
  };
}
