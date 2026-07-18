/**
 * DOM helpers for transient paragraph flashes on the painted layout surface.
 */

// The option shapes live in a DOM-free sibling module so non-browser consumers
// (e.g. the agents package) can type-import them without dragging this file's
// DOM code into their type-check surface. Re-exported here so the public
// `@betteroffice/docx/utils` barrel surface is unchanged.
export type { ParagraphHighlightOptions, ScrollToParaIdOptions } from './paragraphFlashTypes';
import type { ParagraphHighlightOptions } from './paragraphFlashTypes';

/** Default color used by paragraph flashes. */
export const DEFAULT_PARAGRAPH_FLASH_COLOR = 'rgba(255, 235, 59, 0.55)';
/** Default duration for paragraph flashes. */
export const DEFAULT_PARAGRAPH_FLASH_DURATION_MS = 1200;
/** CSS class applied to paragraph fragments during a transient flash. */
export const PARAGRAPH_FLASH_CLASS_NAME = 'docx-paragraph-flash';

const timers = new WeakMap<HTMLElement, ReturnType<typeof setTimeout>>();

function escapeAttributeValue(value: string): string {
  const cssGlobal = globalThis as typeof globalThis & {
    CSS?: { escape?: (value: string) => string };
  };
  if (typeof cssGlobal.CSS?.escape === 'function') {
    return cssGlobal.CSS.escape(value);
  }
  return value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function normalizedColor(options?: ParagraphHighlightOptions): string {
  const color = options?.color?.trim();
  return color || DEFAULT_PARAGRAPH_FLASH_COLOR;
}

function normalizedDurationMs(options?: ParagraphHighlightOptions): number {
  const duration = options?.durationMs;
  if (duration == null) return DEFAULT_PARAGRAPH_FLASH_DURATION_MS;
  if (!Number.isFinite(duration) || duration < 0) return DEFAULT_PARAGRAPH_FLASH_DURATION_MS;
  return duration;
}

/**
 * Find all painted paragraph fragments with a stable `data-para-id`.
 *
 * @public
 */
export function findParagraphFragmentsByParaId(root: ParentNode, paraId: string): HTMLElement[] {
  if (!paraId || !paraId.trim()) return [];
  const escaped = escapeAttributeValue(paraId);
  return Array.from(
    root.querySelectorAll<HTMLElement>(`.layout-paragraph[data-para-id="${escaped}"]`)
  );
}

/**
 * Apply a transient flash to a collection of paragraph elements.
 *
 * @returns the number of elements flashed
 * @public
 */
export function flashParagraphElements(
  elements: Iterable<HTMLElement>,
  options?: ParagraphHighlightOptions
): number {
  let count = 0;
  const color = normalizedColor(options);
  const durationMs = normalizedDurationMs(options);

  for (const el of elements) {
    count++;
    const existingTimer = timers.get(el);
    if (existingTimer !== undefined) clearTimeout(existingTimer);

    el.classList.remove(PARAGRAPH_FLASH_CLASS_NAME);
    // Restart the CSS animation when callers flash the same paragraph twice.
    void el.offsetWidth;
    el.style.setProperty('--docx-paragraph-flash-color', color);
    el.style.setProperty('--docx-paragraph-flash-duration', `${durationMs}ms`);
    el.classList.add(PARAGRAPH_FLASH_CLASS_NAME);

    const timer = setTimeout(() => {
      el.classList.remove(PARAGRAPH_FLASH_CLASS_NAME);
      el.style.removeProperty('--docx-paragraph-flash-color');
      el.style.removeProperty('--docx-paragraph-flash-duration');
      timers.delete(el);
    }, durationMs);
    timers.set(el, timer);
  }

  return count;
}

/**
 * Find paragraph fragments by `paraId` and flash them.
 *
 * @returns whether at least one rendered fragment was found
 * @public
 */
export function flashParagraphFragmentsByParaId(
  root: ParentNode,
  paraId: string,
  options?: ParagraphHighlightOptions
): boolean {
  const fragments = findParagraphFragmentsByParaId(root, paraId);
  return flashParagraphElements(fragments, options) > 0;
}
