import { useCallback, useLayoutEffect, useRef, useState } from 'react';
import type { RefObject } from 'react';

export interface ToolbarOverflowOptions {
  /** Row whose element children are the overflow units, in host order. */
  items: RefObject<HTMLElement | null>;
  /** The "More" trigger, rendered only while something overflows. */
  more: RefObject<HTMLElement | null>;
  /** Whether a unit may move into the overflow menu. */
  canHide(unit: HTMLElement): boolean;
  /** Called when a resize hid the focused element; focus the trigger here. */
  onFocusHidden?(): void;
  /** Width reserved for the trigger before it has been measured. */
  moreWidth?: number;
}

export interface ToolbarOverflowState {
  /** Units currently moved into the overflow menu, in host order. */
  hidden: readonly HTMLElement[];
  /** Re-measures after content changed without a resize. */
  remeasure(): void;
}

const HIDDEN_STYLE = {
  position: 'absolute',
  visibility: 'hidden',
  pointerEvents: 'none',
  insetInlineStart: '0',
} as const satisfies Partial<Record<keyof CSSStyleDeclaration, string>>;

const SCROLL_STYLE = {
  overflowX: 'auto',
  overflowY: 'hidden',
  scrollbarWidth: 'thin',
} as const satisfies Partial<Record<keyof CSSStyleDeclaration, string>>;

type StyleKey = keyof typeof HIDDEN_STYLE | keyof typeof SCROLL_STYLE;
type SavedStyle = Partial<Record<StyleKey, string>>;

const saved = new WeakMap<HTMLElement, SavedStyle>();
const scrolling = new WeakMap<HTMLElement, SavedStyle>();

const CONTROLS = [
  'a[href]',
  'button',
  'input:not([type="hidden"])',
  'select',
  'textarea',
  'summary',
  '[tabindex]',
  '[contenteditable]:not([contenteditable="false"])',
].join(', ');

/**
 * Whether `sources`, elements with overflow-menu entries, cover every control
 * in `unit`, so hiding the unit loses no action.
 */
export function representsUnit(unit: HTMLElement, sources: readonly HTMLElement[]): boolean {
  if (sources.length === 0) return false;
  const controls = Array.from(unit.querySelectorAll(CONTROLS));
  if (unit.matches(CONTROLS)) controls.push(unit);
  return controls.every((control) => sources.some((source) => source.contains(control)));
}

function outerWidth(element: HTMLElement): number {
  const style = getComputedStyle(element);
  const margins = (parseFloat(style.marginLeft) || 0) + (parseFloat(style.marginRight) || 0);
  return element.getBoundingClientRect().width + margins;
}

function innerWidth(element: HTMLElement): number {
  const style = getComputedStyle(element);
  const padding = (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0);
  return element.getBoundingClientRect().width - padding;
}

/** Applies `style` over the element's inline style, or restores it; returns whether it changed. */
function swapStyle(
  element: HTMLElement,
  style: SavedStyle,
  store: WeakMap<HTMLElement, SavedStyle>,
  on: boolean
): boolean {
  const previous = store.get(element);
  if (on === Boolean(previous)) return false;
  const keys = Object.keys(style) as StyleKey[];
  if (on) {
    const values: SavedStyle = {};
    for (const key of keys) {
      values[key] = element.style[key] as string;
      element.style[key] = style[key] ?? '';
    }
    store.set(element, values);
  } else {
    for (const key of keys) element.style[key] = previous?.[key] ?? '';
    store.delete(element);
  }
  return true;
}

function setHidden(element: HTMLElement, hidden: boolean): void {
  if (!swapStyle(element, HIDDEN_STYLE, saved, hidden)) return;
  if (hidden) element.setAttribute('aria-hidden', 'true');
  else element.removeAttribute('aria-hidden');
  (element as HTMLElement & { inert: boolean }).inert = hidden;
}

/** Scrolls `row` by the least amount that shows `target`, leaving other scrollers alone. */
function reveal(row: HTMLElement, target: Element): void {
  const bounds = row.getBoundingClientRect();
  const box = target.getBoundingClientRect();
  if (box.left < bounds.left) row.scrollLeft -= bounds.left - box.left;
  else if (box.right > bounds.right) {
    row.scrollLeft += Math.min(box.right - bounds.right, box.left - bounds.left);
  }
}

function sameUnits(a: readonly HTMLElement[], b: readonly HTMLElement[]): boolean {
  return a.length === b.length && a.every((element, index) => element === b[index]);
}

/**
 * Moves trailing units of a toolbar row into an overflow menu when the row is
 * too narrow. Hidden units stay mounted for measurement, but are inert and out
 * of the accessibility tree. When the units that cannot hide still do not fit,
 * the row scrolls, and a focused control scrolls into view.
 */
export function useToolbarOverflow(options: ToolbarOverflowOptions): ToolbarOverflowState {
  const optionsRef = useRef(options);
  optionsRef.current = options;
  const [hidden, setHiddenUnits] = useState<readonly HTMLElement[]>([]);
  const hiddenRef = useRef(hidden);
  hiddenRef.current = hidden;
  const moreWidthRef = useRef(options.moreWidth ?? 32);

  const measure = useCallback(() => {
    const { items, more, canHide, onFocusHidden } = optionsRef.current;
    const row = items.current;
    const rail = row?.parentElement;
    if (!row || !rail) return;
    const trigger = more.current;
    if (trigger) moreWidthRef.current = outerWidth(trigger);
    const units = Array.from(row.children).filter(
      (child): child is HTMLElement => child instanceof HTMLElement
    );
    const gap = parseFloat(getComputedStyle(row).columnGap) || 0;
    const available = innerWidth(rail);
    const widths = units.map(outerWidth);
    const total =
      widths.reduce((sum, width) => sum + width, 0) + gap * Math.max(0, units.length - 1);
    const next: HTMLElement[] = [];
    let overflows = false;
    if (total > available + 0.5) {
      const budget = available - moreWidthRef.current;
      const hideable = units.map((unit) => canHide(unit));
      let used = units.reduce(
        (sum, _unit, index) => (hideable[index] ? sum : sum + widths[index] + gap),
        0
      );
      let cut = false;
      units.forEach((unit, index) => {
        if (!hideable[index]) return;
        if (!cut && used + widths[index] <= budget) {
          used += widths[index] + gap;
          return;
        }
        cut = true;
        next.push(unit);
      });
      overflows = used - gap > (next.length > 0 ? budget : available) + 0.5;
    }
    swapStyle(row, SCROLL_STYLE, scrolling, overflows);
    for (const unit of units) setHidden(unit, next.includes(unit));
    for (const unit of hiddenRef.current) if (!units.includes(unit)) setHidden(unit, false);
    const active = document.activeElement;
    if (active && next.some((unit) => unit.contains(active))) onFocusHidden?.();
    else if (overflows && active && row.contains(active)) reveal(row, active);
    if (!sameUnits(next, hiddenRef.current)) {
      hiddenRef.current = next;
      setHiddenUnits(next);
    }
  }, []);

  useLayoutEffect(() => {
    measure();
  });

  useLayoutEffect(() => {
    const row = optionsRef.current.items.current;
    if (!row) return;
    const observed = new Set<Element>();
    const observer =
      typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(() => measure());
    const observe = () => {
      if (!observer) return;
      for (const element of [row, row.parentElement, ...Array.from(row.children)]) {
        if (element && !observed.has(element)) {
          observed.add(element);
          observer.observe(element);
        }
      }
    };
    observe();
    const mutations =
      typeof MutationObserver === 'undefined'
        ? null
        : new MutationObserver(() => {
            observe();
            measure();
          });
    mutations?.observe(row, { childList: true });
    const onFocusIn = (event: FocusEvent) => {
      if (event.target instanceof Element) reveal(row, event.target);
    };
    row.addEventListener('focusin', onFocusIn);
    window.addEventListener('resize', measure);
    void document.fonts?.ready.then(measure, () => undefined);
    return () => {
      observer?.disconnect();
      mutations?.disconnect();
      row.removeEventListener('focusin', onFocusIn);
      window.removeEventListener('resize', measure);
      swapStyle(row, SCROLL_STYLE, scrolling, false);
      for (const unit of Array.from(row.children)) {
        if (unit instanceof HTMLElement) setHidden(unit, false);
      }
    };
  }, [measure]);

  return { hidden, remeasure: measure };
}
