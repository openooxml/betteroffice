/**
 * Flow — turn a ProseMirror document into the flat `LayoutBlock[]` the layout
 * pipeline consumes, including demotion of block-like floating tables.
 *
 * @packageDocumentation
 * @public
 */

export * from './toLayoutBlocks';
export { isBlockLikeFloatingTable, demoteBlockLikeFloatingTables } from './floatingTable';
