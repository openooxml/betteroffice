/**
 * Host wiring for the Rust pagination source — the sole pagination engine.
 *
 * Deliberately thinner than `useRustMeasurement`: the source owns the wasm
 * load (started at construction, shared with the measurement engine's
 * module), so the hook only owns identity. Readiness is the layout
 * pipeline's concern: `useLayoutPipeline` defers passes while
 * `source.isReady()` is false and re-runs once `whenReady()` settles —
 * the same defer-and-rerun pattern as the measurement readiness gate.
 */

import { useMemo } from 'react';

import {
  createRustPaginationSource,
  type LayoutPaginationSource,
  type RustLayoutEngine,
} from '@betteroffice/docx/layout/pagination';

/**
 * Build the mandatory Rust pagination source. Identity-stable for the
 * component's lifetime (useLayoutPipeline's dep array holds it).
 */
export function useRustPagination(engine?: RustLayoutEngine | null): LayoutPaginationSource {
  return useMemo(
    () => createRustPaginationSource(engine ? { engine } : undefined),
    [engine]
  );
}
