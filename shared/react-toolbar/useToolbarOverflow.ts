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

type HiddenKey = keyof typeof HIDDEN_STYLE;

const saved = new WeakMap<HTMLElement, Partial<Record<HiddenKey, string>>>();

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

function setHidden(element: HTMLElement, hidden: boolean): void {
  const previous = saved.get(element);
  if (hidden && !previous) {
    const values: Partial<Record<HiddenKey, string>> = {};
    for (const key of Object.keys(HIDDEN_STYLE) as HiddenKey[]) {
      values[key] = element.style[key] as string;
      element.style[key] = HIDDEN_STYLE[key];
    }
    saved.set(element, values);
    element.setAttribute('aria-hidden', 'true');
    (element as HTMLElement & { inert: boolean }).inert = true;
  } else if (!hidden && previous) {
    for (const key of Object.keys(HIDDEN_STYLE) as HiddenKey[]) {
      element.style[key] = previous[key] ?? '';
    }
    saved.delete(element);
    element.removeAttribute('aria-hidden');
    (element as HTMLElement & { inert: boolean }).inert = false;
  }
}

function sameUnits(a: readonly HTMLElement[], b: readonly HTMLElement[]): boolean {
  return a.length === b.length && a.every((element, index) => element === b[index]);
}

/**
 * Moves trailing units of a toolbar row into an overflow menu when the row is
 * too narrow. Hidden units stay mounted for measurement, but are inert and out
 * of the accessibility tree.
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
    }
    for (const unit of units) setHidden(unit, next.includes(unit));
    for (const unit of hiddenRef.current) if (!units.includes(unit)) setHidden(unit, false);
    const active = document.activeElement;
    if (active && next.some((unit) => unit.contains(active))) onFocusHidden?.();
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
    window.addEventListener('resize', measure);
    void document.fonts?.ready.then(measure, () => undefined);
    return () => {
      observer?.disconnect();
      mutations?.disconnect();
      window.removeEventListener('resize', measure);
      for (const unit of Array.from(row.children)) {
        if (unit instanceof HTMLElement) setHidden(unit, false);
      }
    };
  }, [measure]);

  return { hidden, remeasure: measure };
}
