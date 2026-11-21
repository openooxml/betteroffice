// @betteroffice/docx declaration docx/contentControls.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Headless content-control discovery: DOCX bytes in, the document's content controls out. The Rust
 * engine does the reading; no session, editor or DOM is involved, and control ids and anchors
 * address the returned snapshot only.
 */
import type { DocxContentControlQuery, DocxContentControlReadFailure, DocxContentControlsOptions, DocxContentControlsSnapshot } from '../yrs/contentControls';
/** A read the engine refused as data: unusable options or a limit. */
export declare class DocxContentControlsError extends Error {
    readonly failure: DocxContentControlReadFailure;
    constructor(failure: DocxContentControlReadFailure);
}
/**
 * Lists the content controls of DOCX bytes in document order. Throws
 * {@link DocxContentControlsError} for unusable options or an exceeded limit and an `Error` for
 * bytes that are not a readable DOCX.
 */
export declare function listDocxContentControls(bytes: Uint8Array, options?: DocxContentControlsOptions): Promise<DocxContentControlsSnapshot>;
/** The content controls of DOCX bytes that match `query` exactly; none or several. */
export declare function findDocxContentControls(bytes: Uint8Array, query: DocxContentControlQuery, options?: DocxContentControlsOptions): Promise<DocxContentControlsSnapshot>;
