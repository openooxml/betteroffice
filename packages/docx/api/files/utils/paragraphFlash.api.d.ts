// @betteroffice/docx declaration utils/paragraphFlash.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * DOM helpers for transient paragraph flashes on the painted layout surface.
 */
export type { ParagraphHighlightOptions, ScrollToParaIdOptions } from './paragraphFlashTypes';
import type { ParagraphHighlightOptions } from './paragraphFlashTypes';
/** Default color used by paragraph flashes. */
export declare const DEFAULT_PARAGRAPH_FLASH_COLOR = "rgba(255, 235, 59, 0.55)";
/** Default duration for paragraph flashes. */
export declare const DEFAULT_PARAGRAPH_FLASH_DURATION_MS = 1200;
/** CSS class applied to paragraph fragments during a transient flash. */
export declare const PARAGRAPH_FLASH_CLASS_NAME = "docx-paragraph-flash";
/**
 * Find all painted paragraph fragments with a stable `data-para-id`.
 *
 * @public
 */
export declare function findParagraphFragmentsByParaId(root: ParentNode, paraId: string): HTMLElement[];
/**
 * Apply a transient flash to a collection of paragraph elements.
 *
 * @returns the number of elements flashed
 * @public
 */
export declare function flashParagraphElements(elements: Iterable<HTMLElement>, options?: ParagraphHighlightOptions): number;
/**
 * Find paragraph fragments by `paraId` and flash them.
 *
 * @returns whether at least one rendered fragment was found
 * @public
 */
export declare function flashParagraphFragmentsByParaId(root: ParentNode, paraId: string, options?: ParagraphHighlightOptions): boolean;
