// @betteroffice/docx declaration docx/rezip/createEmpty.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Empty DOCX Template
 *
 * Produces a minimal valid DOCX package — Content Types, package + document
 * rels, an empty body with the standard letter-size section, default
 * Calibri 11 styles, and core/app properties. Used as the base for
 * `createDocx()` when the caller has no existing buffer.
 */
/**
 * Create a new empty DOCX file.
 *
 * @returns Promise resolving to minimal DOCX as ArrayBuffer
 */
export declare function createEmptyDocx(): Promise<ArrayBuffer>;
