/** Typed TypeScript boundary for the Rust package writer. */

import { engineProgress } from '../utils/engineProgress';
import type { BlockContent, Document, HeaderFooter, Hyperlink, Image, Run } from '../types/document';
import { visitTrackedControlContent } from '../utils/trackedControlContent';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import type { LegacyParagraphV1, ParagraphPatchV1 } from './paragraphPatch';
import {
  PARAGRAPH_PATCH_WRITER_VERSION,
  paragraphPatchWriterVersionWire,
  preloadParseWasm,
  writeDocxS13WithWarningsWire,
} from './parseWasm';
import { withSessionRevisionIds } from './sessionRevisionIds';
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

/** Main-document body paragraphs a save replaces by source location, guarded by the part digest. */
export interface RustSourceParagraphs {
  partSha256: string;
  paragraphs: Array<{ path: number[]; block: number }>;
}

export interface RustSelectiveSave {
  changedParaIds: Iterable<string>;
  /** Replace only these paragraphs and keep every other part's source bytes. */
  sourceParagraphs?: RustSourceParagraphs;
}

export interface RustSaveResult {
  buffer: ArrayBuffer;
  determinism: RustSaveDeterminism;
  warnings: string[];
}

/**
 * Paragraph IDs a session save applies: IDs for source paragraphs outside
 * the edited stories by part and `w:p` occurrence, and parts written as their
 * source bytes with only these IDs patched in.
 */
export interface RustParagraphIds {
  removedCommentIds?: string[];
  sectionEdits?: {
    onlySections: boolean;
    sha256: string;
    /** `paragraph` is a `w:p` occurrence, or with `bodyParagraphs` the position among the body's own paragraphs. */
    edits: Array<{ paragraph: number | null; bodyParagraphs?: number; donor: number | null; properties: Record<string, unknown> }>;
  };
  assignments: Array<{ part: string; ordinal: number; paraId: string }>;
  patchedParts: Array<{ part: string; paraIds: Array<[number, string]> }>;
  wholeParts?: Array<{ part: string; sha256: string; changed: number[] }>;
  touchedParts?: Array<{ part: string; touched: number[] }>;
  /**
   * Story parts written as their source bytes with only some paragraphs
   * re-serialized: those whose `sourceOrdinal` is in `changed`, or whose
   * written comments, revisions, notes or relationships differ from their
   * source. `paragraphs` lists every `sourceOrdinal` the model holds for the
   * part, `sha256` the source part it addresses, `removed` the source
   * paragraphs it no longer holds, and `inserted`, in writing order, where
   * each paragraph written without a `sourceOrdinal` goes.
   */
  splicedParts?: Array<{
    part: string;
    sha256: string;
    paragraphs: number[];
    changed: number[];
    paragraphPatches?: ParagraphPatchV1[];
    legacyParagraphs?: LegacyParagraphV1[];
    tablePatches?: Array<{ ordinal: number; before: unknown; after: unknown }>;
    inserted?: Array<{ before: number } | { after: number }>;
    removed?: number[];
    formatOnly?: number[];
  }>;
}

function headerFooterSaveEntries(
  entries: Map<string, HeaderFooter> | undefined
): Array<[string, HeaderFooter]> {
  return entries
    ? [...entries].filter(([, part]) => !isWrittenByCanonical(part, entries))
    : [];
}

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export function modelSaveFields(document: Document) {
  const pkg = document.package;
  return {
    document: pkg.document,
    headerEntries: [...(pkg.headers ?? [])],
    footerEntries: [...(pkg.footers ?? [])],
    footnotes: pkg.footnotes,
    endnotes: pkg.endnotes,
    footnoteSeparators: pkg.footnoteSeparators,
    endnoteSeparators: pkg.endnoteSeparators,
    relationshipEntries: [...(pkg.relationships ?? [])],
    numbering: pkg.numbering,
  };
}

const NO_BINDING_CHANGES = 'DOCX_SAVE_BINDINGS 0';

const parsedModels = new WeakMap<Document, {
  fields: ReturnType<typeof modelSaveFields>;
  source: Promise<string | undefined>;
}>();

export function snapshotSaveTree<T>(value: T): T {
  const copies = new WeakMap<object, object>();
  const clone = (field: unknown): unknown => {
    if (!field || typeof field !== 'object') return field;
    const existing = copies.get(field);
    if (existing) return existing;
    if (Array.isArray(field)) {
      const array: unknown[] = [];
      copies.set(field, array);
      for (const child of field) array.push(clone(child));
      return array;
    }
    const object: Record<string, unknown> = Object.create(null);
    copies.set(field, object);
    for (const [key, child] of Object.entries(field)) object[key] = clone(child);
    return object;
  };
  return clone(value) as T;
}

export function sameSaveTree(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    for (let index = 0; index < left.length; index += 1) {
      if (!sameSaveTree(left[index], right[index])) return false;
    }
    return true;
  }
  const a = left as Record<string, unknown>;
  const b = right as Record<string, unknown>;
  const keys = (value: Record<string, unknown>) =>
    Object.keys(value).filter(key => key !== 'sourceOrdinal' && value[key] !== undefined);
  const fields = keys(a);
  return fields.length === keys(b).length &&
    fields.every(key => Object.hasOwn(b, key) && sameSaveTree(a[key], b[key]));
}

export function rememberParsedModel(document: Document): void {
  if (!document.originalBuffer) return;
  parsedModels.set(document, {
    fields: snapshotSaveTree(modelSaveFields(document)),
    source: typeof globalThis.crypto?.subtle !== 'undefined'
      ? sha256Hex(new Uint8Array(document.originalBuffer)).catch(() => undefined)
      : Promise.resolve(undefined),
  });
}

export function parsedModelUnchanged(document: Document): ReturnType<typeof parsedModels.get> {
  const baseline = parsedModels.get(document);
  return baseline && sameSaveTree(modelSaveFields(document), baseline.fields) ? baseline : undefined;
}

/** Encode the public Document through the inverse allowlisted package wire. */
export async function writeDocumentWithRust(
  document: Document,
  originalBuffer: ArrayBuffer,
  options: RustSaveOptions = {},
  selective?: RustSelectiveSave,
  determinism?: RustSaveDeterminism,
  paragraphIds?: RustParagraphIds,
  skipMutations = false,
  yieldSave?: () => Promise<void>,
  styleLanguages?: Readonly<Record<string, unknown>>,
  modelMode?: 'fresh' | 'projected'
): Promise<RustSaveResult> {
  const guarded = modelMode !== 'fresh' && paragraphIds === undefined && selective === undefined;
  await Promise.all([preloadOpcWasm(), preloadParseWasm()]);
  if (yieldSave) await yieldSave();
  for (const part of paragraphIds?.splicedParts ?? []) {
    if (part.changed.length && paragraphPatchWriterVersionWire() !== PARAGRAPH_PATCH_WRITER_VERSION) {
      throw new Error(
        `DOCX_MINIMAL_PATCH [INVALID_PLAN] ${part.part} p#?: ` +
        'paragraphPatchWriterVersion must report PARAGRAPH_PATCH_WRITER_VERSION=1'
      );
    }
  }
  const fixed =
    determinism ??
    ({
      seed: await sha256Hex(new Uint8Array(originalBuffer)),
      now: new Date().toISOString(),
    } satisfies RustSaveDeterminism);
  validateDeterminism(fixed);
  if (yieldSave) await yieldSave();

  const baseline = guarded && modelMode === undefined ? parsedModels.get(document) : undefined;
  const sourceDigest = await baseline?.source;
  const request = {
    ...modelSaveFields(document),
    ...(modelMode === 'fresh' ? { freshPackage: true } : {}),
    ...(modelMode === 'projected' ? { rewriteBody: true } : {}),
    ...(sourceDigest === fixed.seed && parsedModelUnchanged(document)
      ? { unchangedSourceSha256: sourceDigest } : {}),
    determinism: fixed,
    ...(styleLanguages === undefined ? {} : { styleLanguages }),
    ...(!guarded && {
      headerEntries: headerFooterSaveEntries(document.package.headers),
      footerEntries: headerFooterSaveEntries(document.package.footers),
    }),
    options: {
      updateModifiedDate: options.updateModifiedDate ?? true,
      ...(options.modifiedBy === undefined ? {} : { modifiedBy: options.modifiedBy }),
    },
    ...(selective === undefined
      ? {}
      : {
          selective: {
            changedParaIds: [...selective.changedParaIds],
            ...(selective.sourceParagraphs === undefined
              ? {}
              : { sourceParagraphs: selective.sourceParagraphs }),
          },
        }),
    ...(paragraphIds === undefined ? {} : { paragraphIds }),
  };
  engineProgress();
  assertSafeSaveTree(request, 'save');
  engineProgress();
  const wire = withSessionRevisionIds(request, () => JSON.stringify(request));
  engineProgress();
  if (yieldSave) await yieldSave();
  const { bytes, warnings } = writeDocxS13WithWarningsWire(wire, new Uint8Array(originalBuffer));
  engineProgress();
  const buffer = exactArrayBuffer(bytes);
  if (yieldSave) await yieldSave();
  const unchangedBindings = warnings.includes(NO_BINDING_CHANGES);
  if (!selective && !skipMutations && !unchangedBindings) {
    applyRustSaveMutations(document, originalBuffer, buffer);
  }
  return {
    buffer,
    determinism: fixed,
    warnings: [...(document.warnings ?? []), ...warnings.filter(warning => warning !== NO_BINDING_CHANGES)],
  };
}

/** Applies saved relationship IDs to bound model nodes. */
function applyRustSaveMutations(
  document: Document,
  originalBuffer: ArrayBuffer,
  savedBuffer: ArrayBuffer
): void {
  const originalParts = unzipContainer(new Uint8Array(originalBuffer));
  const savedParts = unzipContainer(new Uint8Array(savedBuffer));
  const foldedParts = (parts: typeof originalParts) => new Map(
    Object.entries(parts).map(([path, bytes]) => [path.toLowerCase(), bytes])
  );
  const originalByPath = foldedParts(originalParts);
  const savedByPath = foldedParts(savedParts);

  for (const part of collectParts(document)) {
    const path = part.relsPath.toLowerCase();
    const originalRelationships = partText(originalByPath.get(path)) ?? '';
    const savedRelationships = partText(savedByPath.get(path)) ?? '';
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
    const savedRelationships = partText(savedByPath.get(relsPath.toLowerCase())) ?? '';
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
          for (const inline of content.content) {
            if (inline.type === 'run') visitRun(inline);
            else visitTrackedControlContent(inline, (node) => {
              if (node.type === 'run') visitRun(node);
            }, false, true);
          }
        } else {
          visitTrackedControlContent(content, (node) => {
            if (node.type === 'run') visitRun(node);
          });
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
        if (content.type !== 'run') {
          visitTrackedControlContent(content, (node) => {
            if (node.type === 'hyperlink' && (node.href || node.rId) && !node.anchor) {
              hyperlinks.push(node);
            }
          });
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
  return tag.match(new RegExp(`\\s${name}\\s*=\\s*(["'])(.*?)\\1`))?.[2];
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

/** @internal */
export function assertSafeSaveTree(value: unknown, path: string, active = new WeakSet<object>()): void {
  const segments: Array<string | number> = [];
  const currentPath = (): string => {
    let result = path;
    for (const segment of segments) {
      result += typeof segment === 'number' ? `[${segment}]` : `.${segment}`;
    }
    return result;
  };
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== 'object') return;
    if (active.has(value)) throw new TypeError(`${currentPath()} contains a cycle`);
    active.add(value);
    if (Array.isArray(value)) {
      const length = value.length;
      for (let index = 0; index < length; index += 1) {
        if (!(index in value)) continue;
        segments.push(index);
        visit(value[index]);
        segments.pop();
      }
    } else {
      for (const key of Object.keys(value)) {
        if (FORBIDDEN_KEYS.has(key)) throw new TypeError(`${currentPath()} contains forbidden key ${key}`);
        segments.push(key);
        visit((value as Record<string, unknown>)[key]);
        segments.pop();
      }
    }
    active.delete(value);
  };
  visit(value);
}

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}
