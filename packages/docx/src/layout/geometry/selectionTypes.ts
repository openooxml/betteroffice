/**
 * Selection-overlay geometry value types, shared by the canvas overlays and
 * the display-list queries (`selectionToRectsFromDisplayList` /
 * `getCaretPositionFromDisplayList`).
 *
 * @packageDocumentation
 * @public
 */

/**
 * A rectangle representing part of a selection.
 *
 * @public
 */
export type SelectionRect = {
  /** X coordinate in container space. */
  x: number;
  /** Y coordinate in container space. */
  y: number;
  /** Width of the rectangle. */
  width: number;
  /** Height of the rectangle (typically line height). */
  height: number;
  /** Page index (0-based). */
  pageIndex: number;
};

/**
 * Caret position for collapsed selection.
 *
 * @public
 */
export type CaretPosition = {
  /** X coordinate in container space. */
  x: number;
  /** Y coordinate in container space. */
  y: number;
  /** Height of the caret (line height). */
  height: number;
  /** Page index (0-based). */
  pageIndex: number;
};
