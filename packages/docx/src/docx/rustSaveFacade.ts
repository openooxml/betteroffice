/** Typed TypeScript boundary for the S13 Rust package writer. */

import type { BlockContent, Document, Hyperlink, Image, Run } from '../types/document';
import { sha256Hex } from './canonical/documentCanonical';
import { preloadParseWasm, writeDocxS13Wire } from './parseWasm';
import { collectParts, headerFooterFilename, partText } from './rezip/parts';
import { preloadOpcWasm, unzipContainer } from './wasm';

export interface RustSaveOptions {
  updateModifiedDate?: boolean;
  modifiedBy?: string;
}

export interface RustSaveDeterminism {
  seed: string;
  now: string;
}

export interface RustSelectiveSave {
  changedParaIds: Iterable<string>;
}

export interface RustSaveResult {
  buffer: ArrayBuffer;
  determinism: RustSaveDeterminism;
}

/** Encode the public Document through the inverse allowlisted package wire. */
export async function writeDocumentWithRust(
  document: Document,
  originalBuffer: ArrayBuffer,
  options: RustSaveOptions = {},
  selective?: RustSelectiveSave,
  determinism?: RustSaveDeterminism
): Promise<RustSaveResult> {
  await preloadOpcWasm();
  await preloadParseWasm();
  const fixed =
    determinism ??
    ({
      seed: await sha256Hex(new Uint8Array(originalBuffer)),
      now: new Date().toISOString(),
    } satisfies RustSaveDeterminism);
  validateDeterminism(fixed);

  const pkg = document.package;
  const request = {
    determinism: fixed,
    document: pkg.document,
    headerEntries: [...(pkg.headers?.entries() ?? [])],
    footerEntries: [...(pkg.footers?.entries() ?? [])],
    footnotes: pkg.footnotes ?? [],
    endnotes: pkg.endnotes ?? [],
    footnoteSeparators: pkg.footnoteSeparators ?? [],
    endnoteSeparators: pkg.endnoteSeparators ?? [],
    relationshipEntries: [...(pkg.relationships?.entries() ?? [])],
    ...(pkg.numbering === undefined ? {} : { numbering: pkg.numbering }),
    options: {
      updateModifiedDate: options.updateModifiedDate ?? true,
      ...(options.modifiedBy === undefined ? {} : { modifiedBy: options.modifiedBy }),
    },
    ...(selective === undefined
      ? {}
      : { selective: { changedParaIds: [...selective.changedParaIds] } }),
  };
  assertSafeSaveTree(request, 'save');
  const bytes = writeDocxS13Wire(JSON.stringify(request), new Uint8Array(originalBuffer));
  const buffer = exactArrayBuffer(bytes);
  if (!selective) applyRustSaveMutations(document, originalBuffer, buffer);
  return { buffer, determinism: fixed };
}

/** Preserve the incumbent save contract: newly bound model nodes receive their rIds in-place. */
function applyRustSaveMutations(
  document: Document,
  originalBuffer: ArrayBuffer,
  savedBuffer: ArrayBuffer
): void {
  const originalParts = unzipContainer(new Uint8Array(originalBuffer));
  const savedParts = unzipContainer(new Uint8Array(savedBuffer));

  for (const part of collectParts(document)) {
    const originalRelationships = partText(originalParts[part.relsPath]) ?? '';
    const savedRelationships = partText(savedParts[part.relsPath]) ?? '';
    const originalIds = new Set(
      relationshipTags(originalRelationships).map((tag) => relationshipAttribute(tag, 'Id'))
    );
    const newImageIds = relationshipTags(savedRelationships)
      .filter((tag) => relationshipAttribute(tag, 'Type')?.endsWith('/image'))
      .map((tag) => relationshipAttribute(tag, 'Id'))
      .filter((id): id is string => !!id && !originalIds.has(id));
    const images = collectNewImages(part.blocks);
    if (images.length > newImageIds.length) {
      throw new Error(
        `Rust save image mutation mismatch in ${part.relsPath}: ` +
          `${images.length} model images, ${newImageIds.length} relationships`
      );
    }
    images.forEach((image, index) => {
      image.rId = newImageIds[index];
    });

    const savedHyperlinks = relationshipTags(savedRelationships).filter(
      (tag) =>
        relationshipAttribute(tag, 'Type')?.endsWith('/hyperlink') &&
        relationshipAttribute(tag, 'TargetMode') === 'External'
    );
    for (const hyperlink of collectExternalHyperlinks(part.blocks)) {
      if (!hyperlink.href) {
        if (!relationshipForId(savedRelationships, hyperlink.rId)) hyperlink.rId = undefined;
        continue;
      }
      const match = savedHyperlinks.find(
        (tag) => decodeXmlEntities(relationshipAttribute(tag, 'Target') ?? '') === hyperlink.href
      );
      const id = match && relationshipAttribute(match, 'Id');
      if (id) hyperlink.rId = id;
    }
  }

  const headers = document.package.headers;
  const relationships = document.package.relationships;
  if (!headers || !relationships) return;
  for (const [ownerId, header] of headers) {
    const watermark = header.watermark;
    if (!watermark || watermark.kind !== 'picture') continue;
    const owner = relationships.get(ownerId);
    if (!owner?.target) continue;
    const filename = headerFooterFilename(owner.target).replace(/^word\//, '');
    const relsPath = `word/_rels/${filename}.rels`;
    const savedRelationships = partText(savedParts[relsPath]) ?? '';
    if (relationshipForId(savedRelationships, watermark.relId)) continue;
    const mediaFilename = watermark.mediaPath?.split('/').pop();
    const match = relationshipTags(savedRelationships).find((tag) => {
      if (!relationshipAttribute(tag, 'Type')?.endsWith('/image')) return false;
      const target = relationshipAttribute(tag, 'Target')?.split('/').pop();
      return mediaFilename ? target === mediaFilename : true;
    });
    const id = match && relationshipAttribute(match, 'Id');
    if (id) watermark.relId = id;
  }
}

function collectNewImages(blocks: BlockContent[]): Image[] {
  const images: Image[] = [];
  const visitRun = (run: Run): void => {
    for (const content of run.content) {
      if (
        content.type === 'drawing' &&
        content.image.src?.startsWith('data:') &&
        !content.image.rId
      ) {
        images.push(content.image);
      }
    }
  };
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const content of block.content) {
        if (content.type === 'run') visitRun(content);
        else if (
          content.type === 'insertion' ||
          content.type === 'deletion' ||
          content.type === 'moveFrom' ||
          content.type === 'moveTo'
        ) {
          for (const inline of content.content) if (inline.type === 'run') visitRun(inline);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) images.push(...collectNewImages(cell.content));
      }
    }
  }
  return images;
}

function collectExternalHyperlinks(blocks: BlockContent[]): Hyperlink[] {
  const hyperlinks: Hyperlink[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const content of block.content) {
        if (content.type === 'hyperlink' && (content.href || content.rId) && !content.anchor) {
          hyperlinks.push(content);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) hyperlinks.push(...collectExternalHyperlinks(cell.content));
      }
    } else if (block.type === 'blockSdt') {
      hyperlinks.push(...collectExternalHyperlinks(block.content));
    }
  }
  return hyperlinks;
}

function relationshipTags(xml: string): string[] {
  return [...xml.matchAll(/<Relationship\b[^>]*\/?>/g)].map((match) => match[0]);
}

function relationshipAttribute(tag: string, name: string): string | undefined {
  return tag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
}

function relationshipForId(xml: string, id: string | undefined): string | undefined {
  if (!id) return undefined;
  return relationshipTags(xml).find((tag) => relationshipAttribute(tag, 'Id') === id);
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, decimal: string) => String.fromCodePoint(Number(decimal)))
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, hexadecimal: string) =>
      String.fromCodePoint(parseInt(hexadecimal, 16))
    )
    .replace(/&amp;/g, '&');
}

function validateDeterminism(value: RustSaveDeterminism): void {
  if (!/^[0-9a-f]{64}$/i.test(value.seed)) {
    throw new TypeError('Rust save determinism seed must be a SHA-256 hex digest');
  }
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value.now)) {
    throw new TypeError('Rust save clock must be a UTC ISO-8601 millisecond timestamp');
  }
}

const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

function assertSafeSaveTree(value: unknown, path: string, active = new WeakSet<object>()): void {
  if (value === null || typeof value !== 'object') return;
  if (active.has(value)) throw new TypeError(`${path} contains a cycle`);
  active.add(value);
  if (Array.isArray(value)) {
    value.forEach((entry, index) => assertSafeSaveTree(entry, `${path}[${index}]`, active));
  } else {
    for (const key of Object.keys(value)) {
      if (FORBIDDEN_KEYS.has(key)) throw new TypeError(`${path} contains forbidden key ${key}`);
      assertSafeSaveTree((value as Record<string, unknown>)[key], `${path}.${key}`, active);
    }
  }
  active.delete(value);
}

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}
