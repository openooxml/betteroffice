// @betteroffice/docx declaration docx/parser.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Public DOCX parser facade.
 *
 * Rust owns ZIP/XML parsing. TypeScript only normalizes host inputs,
 * rehydrates the public `Document`, reports progress, and registers fonts.
 * @packageDocumentation
 * @public
 */
import type { Document } from '../types/document';
import type { DocxInput } from '../utils/docxInput';
export type ProgressCallback = (stage: string, percent: number) => void;
export interface ParseOptions {
    onProgress?: ProgressCallback;
    preloadFonts?: boolean;
    parseHeadersFooters?: boolean;
    parseNotes?: boolean;
}
export declare function parseDocx(input: DocxInput, options?: ParseOptions): Promise<Document>;
