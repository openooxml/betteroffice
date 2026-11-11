// @betteroffice/docx declaration docx/rezip.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Public DOCX package I/O wrappers.
 *
 * Model serialization and selective package writing are Rust-owned. The
 * TypeScript kept here is compatibility glue for public ZIP utilities and host
 * values; package inflation/deflation continues through `docx-container` wasm.
 * @packageDocumentation
 * @public
 */
import type { Document } from '../types/document';
import type { RawDocxContent } from './unzip';
import { type PartsMap } from './rezip/parts';
export { findMaxRId } from './rezip/parts';
export type { PartsMap } from './rezip/parts';
export { createEmptyDocx } from './rezip/createEmpty';
export declare const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
export declare const COMMENTS_EXTENDED_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
export declare const COMMENTS_IDS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
export declare const COMMENTS_EXTENSIBLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml";
export interface RepackOptions {
    /** Retained for signature compatibility; `docx-container` chooses compression. */
    compressionLevel?: number;
    updateModifiedDate?: boolean;
    modifiedBy?: string;
}
/** Repack a public Document through the Rust package writer. */
export declare function repackDocx(doc: Document, options?: RepackOptions): Promise<ArrayBuffer>;
export interface RepackResult {
    buffer: ArrayBuffer;
    warnings: string[];
}
/** [`repackDocx`], also returning the non-fatal save diagnostics. */
export declare function repackDocxWithWarnings(doc: Document, options?: RepackOptions): Promise<RepackResult>;
export declare function repackDocxFromRaw(doc: Document, rawContent: RawDocxContent, options?: RepackOptions): Promise<ArrayBuffer>;
export declare function updateDocumentXml(originalBuffer: ArrayBuffer, newDocumentXml: string, options?: RepackOptions): Promise<ArrayBuffer>;
export declare function updateXmlFile(originalBuffer: ArrayBuffer, path: string, content: string, options?: RepackOptions): Promise<ArrayBuffer>;
export declare function updateMultipleFiles(originalBuffer: ArrayBuffer, updates: Map<string, string | ArrayBuffer>, options?: RepackOptions): Promise<ArrayBuffer>;
export declare function applyUpdatesToZip(pkg: PartsMap, updates: Map<string, string | ArrayBuffer>, _options?: RepackOptions): Promise<ArrayBuffer>;
export declare function addRelationship(originalBuffer: ArrayBuffer, relationship: {
    type: string;
    target: string;
    targetMode?: 'External' | 'Internal';
}): Promise<{
    buffer: ArrayBuffer;
    rId: string;
}>;
export declare function addMedia(originalBuffer: ArrayBuffer, filename: string, data: ArrayBuffer, mimeType: string): Promise<{
    buffer: ArrayBuffer;
    rId: string;
    path: string;
}>;
export declare function updateCoreProperties(corePropsXml: string, options: {
    updateModifiedDate?: boolean;
    modifiedBy?: string;
}): string;
export declare function validateDocx(buffer: ArrayBuffer): Promise<{
    valid: boolean;
    errors: string[];
    warnings: string[];
}>;
export declare function isDocxBuffer(buffer: ArrayBuffer): boolean;
export declare function createDocx(doc: Document): Promise<ArrayBuffer>;
/** Compatibility helper for callers that update header/footer parts directly. */
export declare function collectHeaderFooterUpdates(doc: Document): Map<string, string>;
export default repackDocx;
