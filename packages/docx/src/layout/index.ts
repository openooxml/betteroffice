/**
 * Layout — the paged-layout pipeline, one import.
 *
 * Adapter-facing contracts for resident Rust layout, font loading, selection,
 * and browser rendering.
 *
 * @experimental The named exports below are the public contract for adapter
 * authors, but the API is still evolving and may change in minor releases
 * until a third-party adapter validates it.
 * @packageDocumentation
 * @public
 */

// Measurement (the Rust source + float pipeline)
export * from './measure';

// Selection-overlay geometry value types. Hit-testing, click→PM position and
// selection rectangles are Rust display-list queries now
// (`layout/render/displayListQueries.ts`, `layout/render/canvasPointer.ts`);
// only these value shapes remain adapter-facing.
export type { SelectionRect, CaretPosition } from './geometry/selectionTypes';
export type {
  DrawingScene,
  DrawingSceneNode,
  ShapeEffect,
  ShapeFillPaint,
  ShapeStrokePaint,
  ShapeTextBodyProperties,
} from './drawing';

export {
  computeHfCaretRectsFromDisplayList,
  computeHfSelectionRectsFromDisplayList,
  DEFAULT_PAGE_HEIGHT_PX,
  resolveHeaderFooter,
} from './contracts';
export { LayoutSelectionGate } from './selectionGate';
