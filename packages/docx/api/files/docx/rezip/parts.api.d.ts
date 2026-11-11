// @betteroffice/docx declaration docx/rezip/parts.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Part Enumeration
 *
 * Shared helpers for walking the parts of a DOCX package (body, headers,
 * footers, footnotes, endnotes) when registering newly inserted images and
 * hyperlinks against each part's rels file.
 */
import type { Document } from '../../types/document';
import type { BlockContent } from '../../types/content';
/**
 * The assembled DOCX package as a `path -> bytes` map. Both the full-repack and
 * selective-save flows build one of these and hand it to the wasm container.
 */
export type PartsMap = Map<string, Uint8Array>;
export declare function toBytes(content: string | Uint8Array | ArrayBuffer): Uint8Array;
export declare function partText(bytes: Uint8Array | undefined | null): string | undefined;
export declare function rezipPartsToArrayBuffer(pkg: PartsMap): ArrayBuffer;
/**
 * A DOCX part (body, header, or footer) that owns a rels file and may contain
 * newly inserted images/hyperlinks that need to be registered.
 */
export interface Part {
    /** Path to the rels file for this part, e.g. `word/_rels/header1.xml.rels` */
    relsPath: string;
    blocks: BlockContent[];
}
/**
 * Resolve the on-disk filename of a header/footer part from its relationship entry.
 * Returns e.g. `word/header1.xml`.
 */
export declare function headerFooterFilename(target: string): string;
/**
 * Enumerate all parts that may contain newly inserted images/hyperlinks:
 * the document body, every header and footer, and the footnote/endnote parts.
 * Footnotes/endnotes always serialize to the fixed `word/footnotes.xml` /
 * `word/endnotes.xml`, so their rels paths are fixed too.
 */
export declare function collectParts(doc: Document): Part[];
/**
 * Read an existing rels file (or return a minimal stub) and normalize the
 * self-closing form `<Relationships .../>` — which Word emits for empty parts —
 * to the open/close form so our `.replace('</Relationships>', ...)` append works.
 */
export declare function readRelsOrStub(pkg: PartsMap, relsPath: string): string;
/**
 * Find the highest rId number in a relationships XML string.
 */
export declare function findMaxRId(relsXml: string): number;
