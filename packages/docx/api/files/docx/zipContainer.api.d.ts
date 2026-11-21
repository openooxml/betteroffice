// @betteroffice/docx declaration docx/zipContainer.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Zip-container READ facade for the DOCX open path.
 *
 * Backed by the inlined docx-container wasm (see ./wasm): every non-directory
 * entry is inflated in Rust — with the decompression-bomb and path-traversal
 * guards enforced there by construction — and exposed as a normalized, eager
 * read view. Consumers pull parts by path.
 *
 * Read path only. The write/rezip builder assembles a parts map and hands it
 * to the same wasm container via `rezipContainer` (see ./rezip).
 */
/**
 * Eager, normalized read view of an OPC (DOCX) container. All entries are
 * already materialized; lookups are synchronous.
 */
export interface ZipContainerReader {
    paths(): string[];
    file(path: string): Uint8Array | null;
    text(path: string): string | null;
}
export declare function readDocxContainer(buffer: ArrayBuffer): ZipContainerReader;
