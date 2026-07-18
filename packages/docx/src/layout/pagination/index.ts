/**
 * pagination entry point — the pure, DOM-free seam between block
 * measurement and the Rust layout engine.
 *
 * Pagination itself runs in Rust (`crates/docx-layout`, reached through the
 * wasm seam in `./rustPagination`); this barrel carries the plain-data
 * contract that crosses it: the block/extent/fragment type vocabulary
 * (`./types`), the measured-block input shape (`./measuredBlock`), and the
 * DOM-free geometry helpers shared with measurement and the flow pre-pass.
 *
 * @experimental Stable enough for the first-party React adapter, but the
 * API may change in minor releases until a third-party adapter validates
 * it. Pin a version range if you depend on this directly.
 * @packageDocumentation
 * @public
 */

export * from './types';
export * from './measuredBlock';

// Per-section geometry resolution (consumed by layout/regions/sectionGeometry).
export { collectSectionConfigs } from './sectionConfigs';
export type { SectionLayoutConfig } from './sectionConfigs';

// The mandatory Rust pagination source (wasm docx-layout seam).
export {
  createRustPaginationSource,
  getRustLayoutEngine,
  type LayoutPaginationSource,
  type RustLayoutEngine,
  type RustPaginationSource,
  type RustPaginationSourceOptions,
  type RustPaginationStats,
} from './rustPagination';

export type { FootnoteContent } from './types';
export {
  isFloatingTextBoxBlock,
  floatingTextBoxWrapsText,
  floatingTextBoxReservesBand,
  type TextBoxFlowAttrs,
} from './textBoxFlow';
export { isFloatingWrapType, isWrapNone, wrapsAroundText } from '../../docx/wrapTypes';
