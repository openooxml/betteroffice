// @betteroffice/docx declaration layout/index.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
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
export * from './measure';
export type { SelectionRect, CaretPosition } from './geometry/selectionTypes';
export type { DrawingScene, DrawingSceneNode, ShapeEffect, ShapeFillPaint, ShapeStrokePaint, ShapeTextBodyProperties, } from './drawing';
export { DEFAULT_PAGE_HEIGHT_PX, resolveHeaderFooter } from './contracts';
export { LayoutSelectionGate } from './selectionGate';
