import type { PptxFontFace } from '../types';

/** What `renderPng` draws: a slide by zero-based index or by its ID. */
export type PptxPngTarget = { slide: number | string };

export interface PptxPngOptions {
  /** Device pixels per slide pixel. Defaults to 1. */
  scale?: number;
  /** `'slide'` (default) paints white under the slide, as the editor shows it; `'transparent'` nothing; or a `#rrggbb` color. */
  background?: 'slide' | 'transparent' | `#${string}`;
  /** Scratch pixels every shadow on the slide may blur between them. */
  maxShadowPixels?: number;
  signal?: AbortSignal;
  /**
   * Faces to paint with besides those the presentation registered, matched by
   * family and style. One that fails to load rejects with `PngExportError`
   * `font-unavailable`.
   */
  fonts?: ReadonlyArray<PptxFontFace>;
}
