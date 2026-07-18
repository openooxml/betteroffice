/**
 * Option shapes for `scrollToParaId(paraId, { highlight })`.
 *
 * Kept in a DOM-free module (no imports, no browser globals) so non-browser
 * consumers can type-import them without pulling `paragraphFlash`'s DOM code
 * into their type-check surface.
 */

/**
 * Customization for the transient paragraph flash applied by
 * `scrollToParaId(paraId, { highlight })`.
 *
 * @public
 */
export interface ParagraphHighlightOptions {
  /** CSS color used for the transient paragraph flash. Defaults to yellow. */
  color?: string;
  /** How long the flash remains visible before it is removed. Defaults to 1200ms. */
  durationMs?: number;
}

/**
 * Optional reveal behavior for `scrollToParaId`.
 *
 * @public
 */
export interface ScrollToParaIdOptions {
  /** Flash rendered paragraph fragments after scrolling to the paragraph. */
  highlight?: ParagraphHighlightOptions;
}
