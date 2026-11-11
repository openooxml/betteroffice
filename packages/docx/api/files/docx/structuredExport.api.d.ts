// @betteroffice/docx declaration docx/structuredExport.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Headless structured export: DOCX bytes in, read-only structured content or Markdown out. The
 * Rust engine does the reading and anchors address the returned snapshot only. Plain exports
 * involve no session, editor, DOM or font; a paged export lays the bytes out in a private session
 * with exactly the fonts it is given, in a font store no other session shares.
 */
import type { Document } from '../types/document';
import type { YrsRenderEnv } from '../yrs';
import type { DocxHeadlessLayoutOptions, DocxLayoutMap, DocxPageExportOptions, DocxPageMarkdownOptions, DocxPagedStructuredContent, DocxSnapshotLayoutMap } from '../yrs/pagedExport';
import type { DocxExportFailure, DocxExportOptions, DocxMarkdownContent, DocxMarkdownOptions, DocxStructuredContent } from '../yrs/structuredExport';
/** An export the engine refused as data: unusable options. */
export declare class DocxExportError extends Error {
    readonly failure: DocxExportFailure;
    constructor(failure: DocxExportFailure);
}
/**
 * Exports DOCX bytes as structured content. Throws {@link DocxExportError} for unusable options
 * and an `Error` for bytes that are not a readable DOCX.
 */
export declare function exportDocxStructured(bytes: Uint8Array, options: DocxExportOptions): Promise<DocxStructuredContent>;
/** {@link exportDocxStructured} rendered as Markdown with source anchors. */
export declare function exportDocxMarkdown(bytes: Uint8Array, options: DocxExportOptions): Promise<DocxMarkdownContent>;
/** Renders structured content (schema version 1) as Markdown. */
export declare function renderDocxMarkdown(content: DocxStructuredContent, options?: DocxMarkdownOptions): Promise<DocxMarkdownContent>;
/**
 * The render environment the editor lays `document` out with, with `overrides` applied.
 * @internal
 */
export declare function renderEnvironment(document: Document, overrides: DocxHeadlessLayoutOptions['renderEnvironment']): YrsRenderEnv;
/**
 * Exports DOCX bytes as structured content with a page map. A private session opens the bytes
 * and lays every section, header, footer and note out with exactly `layout.fonts`, registered in
 * a measurement font store of its own so no editor's fonts change, and exports from that one
 * captured state. Anchors address the returned snapshot and the map is deterministic for the
 * same bytes, fonts and options. Throws {@link DocxExportError} for a refusal, a `TypeError` for
 * an unknown, repeated or undecodable font key and an `Error` for bytes that are not a readable
 * DOCX or fonts the engine rejects.
 */
export declare function exportDocxStructuredWithPages(bytes: Uint8Array, options: DocxPageExportOptions, layout: DocxHeadlessLayoutOptions): Promise<DocxPagedStructuredContent<DocxSnapshotLayoutMap>>;
/**
 * Renders a paged export as Markdown. With `pageMarkers`, each block marker is followed by
 * `<!-- docx-pages: N=label ... -->`: the physical page indexes and percent-encoded displayed
 * labels of every page showing the block. No page break is written into the text. Throws
 * {@link DocxExportError} when the page map was not captured with the content.
 */
export declare function renderDocxMarkdownWithPages(content: DocxPagedStructuredContent<DocxLayoutMap | DocxSnapshotLayoutMap>, options?: DocxPageMarkdownOptions): Promise<DocxMarkdownContent>;
