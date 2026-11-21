// @betteroffice/docx declaration utils/docxInput.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Flexible input types for DOCX documents.
 *
 * Accepts any common binary format so consumers don't need to manually
 * convert before passing data to the editor or parser.
 */
/**
 * Any binary representation of a DOCX file that the editor can consume.
 *
 * - `ArrayBuffer` — from `FileReader.readAsArrayBuffer()` or `fetch().arrayBuffer()`
 * - `Uint8Array` — from Node.js `fs.readFile()` or streaming APIs
 * - `Blob` — from drag-and-drop or `<input type="file">`
 * - `File` — subclass of Blob, from `<input type="file">`
 */
export type DocxInput = ArrayBuffer | Uint8Array | Blob | File;
/**
 * Normalize any {@link DocxInput} into an `ArrayBuffer` for internal use.
 */
export declare function toArrayBuffer(input: DocxInput): Promise<ArrayBuffer>;
