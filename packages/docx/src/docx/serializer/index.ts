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

import type {
  BlockContent,
  Comment,
  Document,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  Paragraph,
  Run,
  SectionProperties,
  Table,
} from '../../types/document';
import {
  serializeDocxS10Wire,
  serializeDocxS11Wire,
  serializeDocxS12Wire,
} from '../parseWasm';

const STANDALONE_DETERMINISM = {
  seed: '0'.repeat(64),
  now: '2000-01-01T00:00:00.000Z',
};

type SerializerStage = 'S10' | 'S11' | 'S12';

function serializeWithRust(
  stage: SerializerStage,
  request: Record<string, unknown>
): string {
  const input = JSON.stringify({ ...request, determinism: STANDALONE_DETERMINISM });
  const responseJson =
    stage === 'S10'
      ? serializeDocxS10Wire(input)
      : stage === 'S11'
        ? serializeDocxS11Wire(input)
        : serializeDocxS12Wire(input);
  const response = JSON.parse(responseJson) as unknown;
  if (!response || typeof response !== 'object') {
    throw new TypeError(`${stage} Rust serializer returned a non-object response`);
  }
  const value = response as { wireVersion?: unknown; family?: unknown; xml?: unknown };
  if (
    value.wireVersion !== 1 ||
    value.family !== request.family ||
    typeof value.xml !== 'string'
  ) {
    throw new TypeError(`invalid ${stage} Rust serializer response for ${String(request.family)}`);
  }
  return value.xml;
}

function serializeContent(block: BlockContent): string {
  switch (block.type) {
    case 'paragraph':
      return serializeParagraph(block);
    case 'table':
      return serializeTable(block);
    case 'blockSdt':
      return serializeWithRust('S11', { family: 'blockSdt', sdt: block });
  }
}

/** Serialize a single block content item. */
export function serializeBlockContent(block: BlockContent): string {
  return serializeContent(block);
}

/** Serialize a complete `word/document.xml` part. */
export function serializeDocument(doc: Document): string {
  return serializeWithRust('S12', { family: 'document', body: doc.package.document });
}

/** Serialize the contents of `w:body`, without the body tags. */
export function serializeDocumentBody(body: DocumentBody): string {
  const documentXml = serializeWithRust('S12', { family: 'document', body });
  const startTag = '<w:body>';
  const endTag = '</w:body>';
  const start = documentXml.indexOf(startTag);
  const end = documentXml.lastIndexOf(endTag);
  if (start < 0 || end < start) throw new Error('Rust document serializer omitted w:body');
  return documentXml.slice(start + startTag.length, end);
}

/** Serialize section properties. */
export function serializeSectionProperties(props: SectionProperties | undefined): string {
  return serializeWithRust('S10', { family: 'section', properties: props });
}

/** Serialize one paragraph. */
export function serializeParagraph(paragraph: Paragraph): string {
  return serializeWithRust('S11', { family: 'paragraph', paragraph });
}

/** Serialize one run. */
export function serializeRun(run: Run): string {
  return serializeWithRust('S11', { family: 'run', run });
}

/** Serialize one table. */
export function serializeTable(table: Table): string {
  return serializeWithRust('S11', { family: 'table', table });
}

/** Serialize a header or footer part. */
export function serializeHeaderFooter(hf: HeaderFooter): string {
  return serializeWithRust('S12', { family: 'headerFooter', story: hf });
}

/** Serialize the comments part. */
export function serializeComments(comments: Comment[]): string {
  return serializeWithRust('S12', { family: 'comments', comments });
}

/** Serialize the footnotes part. */
export function serializeFootnotes(footnotes: Footnote[]): string {
  return serializeWithRust('S12', { family: 'footnotes', notes: footnotes });
}

/** Serialize the endnotes part. */
export function serializeEndnotes(endnotes: Endnote[]): string {
  return serializeWithRust('S12', { family: 'endnotes', notes: endnotes });
}
