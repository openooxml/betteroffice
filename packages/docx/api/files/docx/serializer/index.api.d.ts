// @betteroffice/docx declaration docx/serializer/index.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * DOCX XML serializers.
 *
 * The public synchronous serializer surface is retained for compatibility, but
 * all OOXML emission is owned by the Rust writer. The fixed context only
 * supplies deterministic IDs to standalone XML calls; package saves use the
 * source-package hash and clock through {@link ../rustSaveFacade}.
 * @packageDocumentation
 * @public
 */
import type { BlockContent, Comment, Document, DocumentBody, Endnote, Footnote, HeaderFooter, Paragraph, Run, SectionProperties, Table } from '../../types/document';
/** Serialize a single block content item. */
export declare function serializeBlockContent(block: BlockContent): string;
/** Serialize a complete `word/document.xml` part. */
export declare function serializeDocument(doc: Document): string;
/** Serialize the contents of `w:body`, without the body tags. */
export declare function serializeDocumentBody(body: DocumentBody): string;
/** Serialize section properties. */
export declare function serializeSectionProperties(props: SectionProperties | undefined): string;
/** Serialize one paragraph. */
export declare function serializeParagraph(paragraph: Paragraph): string;
/** Serialize one run. */
export declare function serializeRun(run: Run): string;
/** Serialize one table. */
export declare function serializeTable(table: Table): string;
/** Serialize a header or footer part. */
export declare function serializeHeaderFooter(hf: HeaderFooter): string;
/** Serialize the comments part. */
export declare function serializeComments(comments: Comment[]): string;
/** Serialize the footnotes part. */
export declare function serializeFootnotes(footnotes: Footnote[]): string;
/** Serialize the endnotes part. */
export declare function serializeEndnotes(endnotes: Endnote[]): string;
