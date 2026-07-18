/**
 * Public DOCX package I/O wrappers.
 *
 * Model serialization and selective package writing are Rust-owned. The
 * TypeScript kept here is compatibility glue for public ZIP utilities and host
 * values; package inflation/deflation continues through `docx-container` wasm.
 * @packageDocumentation
 * @public
 */

import type { HeaderFooter } from '../types/content';
import type { Document } from '../types/document';
import type { RawDocxContent } from './unzip';
import { writeDocumentWithRust } from './rustSaveFacade';
import { serializeHeaderFooter } from './serializer';
import { unzipContainer } from './wasm';
import { createEmptyDocx } from './rezip/createEmpty';
import {
  findMaxRId,
  headerFooterFilename,
  partText,
  rezipPartsToArrayBuffer,
  toBytes,
  type PartsMap,
} from './rezip/parts';

export { findMaxRId } from './rezip/parts';
export type { PartsMap } from './rezip/parts';
export { createEmptyDocx } from './rezip/createEmpty';

export const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
export const COMMENTS_EXTENDED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
export const COMMENTS_IDS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml';
export const COMMENTS_EXTENSIBLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml';

const HEADER_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const FOOTER_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

export interface RepackOptions {
  /** Retained for signature compatibility; `docx-container` chooses compression. */
  compressionLevel?: number;
  updateModifiedDate?: boolean;
  modifiedBy?: string;
}

/** Repack a public Document through the live Rust S13 writer. */
export async function repackDocx(doc: Document, options: RepackOptions = {}): Promise<ArrayBuffer> {
  if (!doc.originalBuffer) {
    throw new Error(
      'Cannot repack document: no original buffer for round-trip. Use createDocx() for new documents.'
    );
  }
  return (await writeDocumentWithRust(doc, doc.originalBuffer, options)).buffer;
}

export async function repackDocxFromRaw(
  doc: Document,
  rawContent: RawDocxContent,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  return repackDocx({ ...doc, originalBuffer: rawContent.originalBuffer }, options);
}

function loadParts(originalBuffer: ArrayBuffer): PartsMap {
  return new Map(Object.entries(unzipContainer(new Uint8Array(originalBuffer))));
}

export async function updateDocumentXml(
  originalBuffer: ArrayBuffer,
  newDocumentXml: string,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  return updateMultipleFiles(
    originalBuffer,
    new Map([['word/document.xml', newDocumentXml]]),
    options
  );
}

export async function updateXmlFile(
  originalBuffer: ArrayBuffer,
  path: string,
  content: string,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  return updateMultipleFiles(originalBuffer, new Map([[path, content]]), options);
}

export async function updateMultipleFiles(
  originalBuffer: ArrayBuffer,
  updates: Map<string, string | ArrayBuffer>,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  return applyUpdatesToZip(loadParts(originalBuffer), updates, options);
}

export async function applyUpdatesToZip(
  pkg: PartsMap,
  updates: Map<string, string | ArrayBuffer>,
  _options: RepackOptions = {}
): Promise<ArrayBuffer> {
  for (const [path, content] of updates) pkg.set(path, toBytes(content));
  return rezipPartsToArrayBuffer(pkg);
}

function addRelationshipToParts(
  pkg: PartsMap,
  relationship: { type: string; target: string; targetMode?: 'External' | 'Internal' }
): string {
  const relsPath = 'word/_rels/document.xml.rels';
  const relsXml = partText(pkg.get(relsPath));
  if (relsXml === undefined) throw new Error('document.xml.rels not found in DOCX');

  const rId = `rId${findMaxRId(relsXml) + 1}`;
  const targetMode = relationship.targetMode === 'External' ? ' TargetMode="External"' : '';
  const node =
    `<Relationship Id="${rId}" Type="${escapeXml(relationship.type)}" ` +
    `Target="${escapeXml(relationship.target)}"${targetMode}/>`;
  const normalized = relsXml.replace(
    /<Relationships([^>]*)\/>/,
    '<Relationships$1></Relationships>'
  );
  pkg.set(relsPath, toBytes(normalized.replace('</Relationships>', `${node}</Relationships>`)));
  return rId;
}

export async function addRelationship(
  originalBuffer: ArrayBuffer,
  relationship: {
    type: string;
    target: string;
    targetMode?: 'External' | 'Internal';
  }
): Promise<{ buffer: ArrayBuffer; rId: string }> {
  const pkg = loadParts(originalBuffer);
  const rId = addRelationshipToParts(pkg, relationship);
  return { buffer: rezipPartsToArrayBuffer(pkg), rId };
}

export async function addMedia(
  originalBuffer: ArrayBuffer,
  filename: string,
  data: ArrayBuffer,
  mimeType: string
): Promise<{ buffer: ArrayBuffer; rId: string; path: string }> {
  const pkg = loadParts(originalBuffer);
  const path = `word/media/${filename}`;
  pkg.set(path, toBytes(data));
  const rId = addRelationshipToParts(pkg, {
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    target: `media/${filename}`,
  });

  const contentTypes = partText(pkg.get('[Content_Types].xml'));
  const extension = filename.split('.').pop()?.toLowerCase() ?? '';
  if (contentTypes !== undefined && extension && !contentTypes.includes(`Extension="${extension}"`)) {
    const node = `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(
      mimeType || contentTypeForExtension(extension)
    )}"/>`;
    pkg.set('[Content_Types].xml', toBytes(contentTypes.replace('</Types>', `${node}</Types>`)));
  }
  return { buffer: rezipPartsToArrayBuffer(pkg), rId, path };
}

function contentTypeForExtension(extension: string): string {
  return (
    {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      tif: 'image/tiff',
      tiff: 'image/tiff',
      svg: 'image/svg+xml',
      webp: 'image/webp',
      wmf: 'image/x-wmf',
      emf: 'image/x-emf',
    } as Record<string, string>
  )[extension] ?? 'application/octet-stream';
}

export function updateCoreProperties(
  corePropsXml: string,
  options: { updateModifiedDate?: boolean; modifiedBy?: string }
): string {
  let result = corePropsXml;
  if (options.updateModifiedDate) {
    const modified =
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}` +
      '</dcterms:modified>';
    result = result.includes('<dcterms:modified')
      ? result.replace(/<dcterms:modified[^<>]*>[^<]*<\/dcterms:modified>/, modified)
      : result.replace('</cp:coreProperties>', `${modified}</cp:coreProperties>`);
  }
  if (options.modifiedBy) {
    const modifier = `<cp:lastModifiedBy>${escapeXml(options.modifiedBy)}</cp:lastModifiedBy>`;
    result = result.includes('<cp:lastModifiedBy')
      ? result.replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, modifier)
      : result.replace('</cp:coreProperties>', `${modifier}</cp:coreProperties>`);
  }
  return result;
}

export async function validateDocx(buffer: ArrayBuffer): Promise<{
  valid: boolean;
  errors: string[];
  warnings: string[];
}> {
  const errors: string[] = [];
  const warnings: string[] = [];
  try {
    const pkg = loadParts(buffer);
    for (const file of ['[Content_Types].xml', 'word/document.xml']) {
      if (!pkg.has(file)) errors.push(`Missing required file: ${file}`);
    }
    for (const file of ['_rels/.rels', 'word/_rels/document.xml.rels', 'word/styles.xml']) {
      if (!pkg.has(file)) warnings.push(`Missing recommended file: ${file}`);
    }
    const documentXml = partText(pkg.get('word/document.xml'));
    if (documentXml !== undefined) {
      if (!documentXml.includes('<?xml')) warnings.push('document.xml missing XML declaration');
      if (!documentXml.includes('<w:document')) errors.push('document.xml missing w:document element');
      if (!documentXml.includes('<w:body>')) errors.push('document.xml missing w:body element');
    }
    const contentTypes = partText(pkg.get('[Content_Types].xml'));
    if (
      contentTypes !== undefined &&
      !contentTypes.includes('word/document.xml') &&
      !contentTypes.includes(
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
      )
    ) {
      warnings.push('Content_Types.xml may be missing document.xml type declaration');
    }
  } catch (error) {
    const message =
      error instanceof Error ? error.message : typeof error === 'string' ? error : 'Unknown error';
    errors.push(`Failed to read as ZIP: ${message}`);
  }
  return { valid: errors.length === 0, errors, warnings };
}

export function isDocxBuffer(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < 4) return false;
  const bytes = new Uint8Array(buffer);
  return bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export async function createDocx(doc: Document): Promise<ArrayBuffer> {
  const originalBuffer = await createEmptyDocx();
  return repackDocx({ ...doc, originalBuffer });
}

/** Compatibility helper for callers that update header/footer parts directly. */
export function collectHeaderFooterUpdates(doc: Document): Map<string, string> {
  const updates = new Map<string, string>();
  const relationships = doc.package.relationships;
  if (!relationships) return updates;
  const stories: Array<{
    map: Map<string, HeaderFooter> | undefined;
    relationshipType: string;
  }> = [
    { map: doc.package.headers, relationshipType: HEADER_RELATIONSHIP },
    { map: doc.package.footers, relationshipType: FOOTER_RELATIONSHIP },
  ];
  for (const { map, relationshipType } of stories) {
    for (const [rId, story] of map ?? []) {
      const relationship = relationships.get(rId);
      if (relationship?.type === relationshipType && relationship.target) {
        updates.set(headerFooterFilename(relationship.target), serializeHeaderFooter(story));
      }
    }
  }
  return updates;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export default repackDocx;
