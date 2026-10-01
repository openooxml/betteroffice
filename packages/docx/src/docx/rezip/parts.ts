/**
 * Part Enumeration
 *
 * Shared helpers for walking the parts of a DOCX package (body, headers,
 * footers, footnotes, endnotes) when registering newly inserted images and
 * hyperlinks against each part's rels file.
 */

import type { Document } from '../../types/document';
import type { BlockContent, HeaderFooter } from '../../types/content';
import { RELATIONSHIP_TYPES } from '../relsParser';
import { rezipContainer } from '../wasm';

/**
 * The assembled DOCX package as a `path -> bytes` map. Both the full-repack and
 * selective-save flows build one of these and hand it to the wasm container.
 */
export type PartsMap = Map<string, Uint8Array>;

const utf8Encoder = new TextEncoder();
// ignoreBOM keeps a leading BOM in the decoded string so a read→modify→write
// round-trips a part unchanged (the default TextDecoder would strip it).
const utf8Decoder = new TextDecoder('utf-8', { ignoreBOM: true });

// normalize part content to bytes. strings become utf-8 with no BOM.
export function toBytes(content: string | Uint8Array | ArrayBuffer): Uint8Array {
  if (typeof content === 'string') return utf8Encoder.encode(content);
  if (content instanceof Uint8Array) return content;
  return new Uint8Array(content);
}

// decode a part's bytes as utf-8 text, or undefined when the part is absent.
export function partText(bytes: Uint8Array | undefined | null): string | undefined {
  return bytes ? utf8Decoder.decode(bytes) : undefined;
}

// deflate the assembled parts map into a standalone DOCX ArrayBuffer. copies the
// bytes out of wasm memory into an exact-length buffer.
export function rezipPartsToArrayBuffer(pkg: PartsMap): ArrayBuffer {
  const bytes = rezipContainer(Object.fromEntries(pkg));
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
}

/**
 * A DOCX part (body, header, or footer) that owns a rels file and may contain
 * newly inserted images/hyperlinks that need to be registered.
 */
export interface Part {
  /** Path to the rels file for this part, e.g. `word/_rels/header1.xml.rels` */
  relsPath: string;
  blocks: BlockContent[];
}

const EMPTY_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

/**
 * Resolve the on-disk filename of a header/footer part from its relationship entry.
 * Returns e.g. `word/header1.xml`.
 */
export function headerFooterFilename(target: string): string {
  return target.startsWith('/') ? target.slice(1) : `word/${target}`;
}

/**
 * Enumerate all parts that may contain newly inserted images/hyperlinks:
 * the document body, every header and footer, and the footnote/endnote parts.
 * Footnotes/endnotes always serialize to the fixed `word/footnotes.xml` /
 * `word/endnotes.xml`, so their rels paths are fixed too.
 */
export function collectParts(doc: Document): Part[] {
  const parts: Part[] = [
    { relsPath: 'word/_rels/document.xml.rels', blocks: doc.package.document.content },
  ];

  const noteBlocks = (notes: { content: BlockContent[] }[] | undefined): BlockContent[] =>
    (notes ?? []).flatMap((note) => note.content);

  const footnoteBlocks = [
    ...noteBlocks(doc.package.footnoteSeparators),
    ...noteBlocks(doc.package.footnotes),
  ];
  if (footnoteBlocks.length > 0) {
    parts.push({ relsPath: 'word/_rels/footnotes.xml.rels', blocks: footnoteBlocks });
  }

  const endnoteBlocks = [
    ...noteBlocks(doc.package.endnoteSeparators),
    ...noteBlocks(doc.package.endnotes),
  ];
  if (endnoteBlocks.length > 0) {
    parts.push({ relsPath: 'word/_rels/endnotes.xml.rels', blocks: endnoteBlocks });
  }

  const rels = doc.package.relationships;
  if (!rels) return parts;

  const addHeaderFooterParts = (map: Map<string, HeaderFooter> | undefined, type: string) => {
    if (!map) return;
    for (const [rId, hf] of map.entries()) {
      const rel = rels.get(rId);
      if (!rel || rel.type !== type || !rel.target) continue;
      const filename = headerFooterFilename(rel.target);
      const basename = filename.replace(/^word\//, '');
      parts.push({ relsPath: `word/_rels/${basename}.rels`, blocks: hf.content });
    }
  };

  addHeaderFooterParts(doc.package.headers, RELATIONSHIP_TYPES.header);
  addHeaderFooterParts(doc.package.footers, RELATIONSHIP_TYPES.footer);

  return parts;
}

/**
 * Read an existing rels file (or return a minimal stub) and normalize the
 * self-closing form `<Relationships .../>` — which Word emits for empty parts —
 * to the open/close form so our `.replace('</Relationships>', ...)` append works.
 */
export function readRelsOrStub(pkg: PartsMap, relsPath: string): string {
  const xml = partText(pkg.get(relsPath)) ?? EMPTY_RELS_XML;
  return xml.replace(/<Relationships([^>]*)\/>/, '<Relationships$1></Relationships>');
}

/**
 * Find the highest rId number in a relationships XML string.
 */
export function findMaxRId(relsXml: string): number {
  let maxId = 0;
  for (const match of relsXml.matchAll(/Id="rId(\d+)"/g)) {
    const id = parseInt(match[1], 10);
    if (id > maxId) maxId = id;
  }
  return maxId;
}
