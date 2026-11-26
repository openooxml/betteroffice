import type { Document, HeaderFooter } from '../types/document';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import { parseDocumentWithRust } from './rustParseFacade';
import { isSessionRevisionInfo } from './sessionRevisionIds';
import { unzipContainer } from './wasm';

const MODEL_PARTS_REWRITTEN = 'DOCX_MODEL_PARTS_REWRITTEN';
const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

function ownValue<T extends object, K extends keyof T>(object: T, key: K): T[K] {
  const descriptor = Object.getOwnPropertyDescriptor(object, key);
  if (descriptor && 'value' in descriptor) return descriptor.value;
  if (!descriptor && !(key in object)) return undefined as T[K];
  throw new TypeError('Cannot compare an accessor or inherited save value');
}

function mapEntries<K, V>(entries: Map<K, V> | undefined): Array<[K, V]> {
  if (!entries) return [];
  if (Object.getPrototypeOf(entries) !== Map.prototype || Reflect.ownKeys(entries).length) {
    throw new TypeError('Cannot compare custom save maps');
  }
  return [...entries];
}

function headerFooterSaveEntries(
  entries: Map<string, HeaderFooter> | undefined
): Array<[string, HeaderFooter]> {
  return entries
    ? mapEntries(entries).filter(([, part]) => !isWrittenByCanonical(part, entries))
    : [];
}

function snapshot(value: unknown, active: WeakSet<object>): unknown {
  if (value === null || typeof value !== 'object') {
    if (
      value === null ||
      value === undefined ||
      typeof value === 'string' ||
      typeof value === 'boolean' ||
      (typeof value === 'number' && Number.isFinite(value) && !Object.is(value, -0))
    ) return value;
    throw new TypeError('Cannot compare a non-JSON save value');
  }
  if (active.has(value) || isSessionRevisionInfo(value) || 'toJSON' in value) {
    throw new TypeError('Cannot compare a cycle, session revision or custom JSON value');
  }
  const array = Array.isArray(value);
  if (Object.getPrototypeOf(value) !== (array ? Array.prototype : Object.prototype)) {
    throw new TypeError('Cannot compare custom save objects');
  }
  active.add(value);
  try {
    if (array) {
      return Array.from(
        { length: ownValue(value, 'length') },
        (_, index) => snapshot(ownValue(value, index), active)
      );
    }
    const record = value as Record<string, unknown>;
    const copy: Record<string, unknown> = Object.create(null);
    for (const key of Object.keys(record)) {
      if (FORBIDDEN_KEYS.has(key)) throw new TypeError('Cannot compare a forbidden save key');
      copy[key] = snapshot(ownValue(record, key), active);
    }
    return copy;
  } finally {
    active.delete(value);
  }
}

/** Serializes the fields the model writer accepts from a plain document. */
export function serializeDocumentSaveModel(document: Document): string {
  const pkg = ownValue(document, 'package');
  const numbering = ownValue(pkg, 'numbering');
  const model = {
    document: ownValue(pkg, 'document'),
    headerEntries: headerFooterSaveEntries(ownValue(pkg, 'headers')),
    footerEntries: headerFooterSaveEntries(ownValue(pkg, 'footers')),
    footnotes: ownValue(pkg, 'footnotes') ?? [],
    endnotes: ownValue(pkg, 'endnotes') ?? [],
    footnoteSeparators: ownValue(pkg, 'footnoteSeparators') ?? [],
    endnoteSeparators: ownValue(pkg, 'endnoteSeparators') ?? [],
    relationshipEntries: mapEntries(ownValue(pkg, 'relationships')),
    ...(numbering === undefined ? {} : { numbering }),
  };
  return JSON.stringify(snapshot(model, new WeakSet()));
}

export function modelMatchesSource(doc: Document, source: ArrayBuffer): boolean {
  try {
    const { document } = parseDocumentWithRust(source, {
      parseHeadersFooters: true,
      parseNotes: true,
      detectVariables: false,
      includeCanonical: false,
    });
    return serializeDocumentSaveModel(doc) === serializeDocumentSaveModel(document);
  } catch {
    return false;
  }
}

export function modelPartsRewrittenWarning(source: ArrayBuffer, output: ArrayBuffer): string {
  let parts: string;
  try {
    const original = unzipContainer(new Uint8Array(source));
    const written = unzipContainer(new Uint8Array(output));
    const paths = Object.entries(written)
      .filter(([path, bytes]) => {
        const before = original[path];
        return path !== 'docProps/core.xml' && (!before ||
          before.byteLength !== bytes.byteLength || bytes.some((byte, i) => byte !== before[i]));
      })
      .map(([path]) => path)
      .sort();
    parts = JSON.stringify(paths);
  } catch {
    parts = '(part names unavailable)';
  }
  return `${MODEL_PARTS_REWRITTEN} ${parts} Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`;
}
