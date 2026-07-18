/**
 * Shared types and primitive helpers used across toLayoutBlocks sub-modules.
 */

import type { Theme } from '../../../types/document';

/**
 * Options for the conversion.
 */
export type ToLayoutBlocksOptions = {
  /** Default font family. */
  defaultFont?: string;
  /** Default font size in points. */
  defaultSize?: number;
  /** Theme for resolving theme colors. */
  theme?: Theme | null;
  /** Page content height in pixels (pageHeight - marginTop - marginBottom). Images taller than this are scaled down to fit. */
  pageContentHeight?: number;
  /**
   * Document-wide `w:defaultTabStop` (§17.6.13) in twips. Stamped onto each
   * paragraph's attrs so the layout-time list-marker helper can snap body
   * text to the default tab grid when no custom `w:tabs` are defined.
   * Default 720 twips (Word's spec default).
   */
  defaultTabStopTwips?: number;
  /**
   * @internal Allocated by toLayoutBlocks() and threaded through table /
   * text-box conversion so list numbering stays continuous across containers.
   * Keyed by abstractNumId when known (ECMA-376 §17.9.18: numIds sharing one
   * abstractNum share counter state); falls back to numId.
   */
  listCounters?: Map<number, number[]>;
  /**
   * @internal Tracks `${numId}:${ilvl}` pairs whose startOverride has already
   * been applied. Per ECMA-376 §17.9.27 the override fires the first time
   * each level of a numId is encountered, so a numId with overrides on
   * multiple ilvls fires each one independently.
   */
  listSeenNumIds?: Set<string>;
};

/**
 * Convert twips to pixels (1 twip = 1/1440 inch, 1 inch = 96 CSS px).
 * No rounding — precision prevents cumulative layout drift across paragraphs.
 */
export function twipsToPixels(twips: number): number {
  return (twips / 1440) * 96;
}

/**
 * Constrain image dimensions to fit within the page content area.
 * Scales proportionally if height exceeds pageContentHeight.
 */
export function constrainImageToPage(
  width: number,
  height: number,
  pageContentHeight: number | undefined
): { width: number; height: number } {
  if (!pageContentHeight || height <= pageContentHeight) {
    return { width, height };
  }
  const scale = pageContentHeight / height;
  return { width: Math.round(width * scale), height: pageContentHeight };
}

let blockIdCounter = 0;

// hand out monotonically increasing ids for synthesized layout items
export function nextBlockId(): string {
  return `block-${++blockIdCounter}`;
}

// rewind the id sequence so test runs start from a known value
export function resetBlockIdCounter(): void {
  blockIdCounter = 0;
}
