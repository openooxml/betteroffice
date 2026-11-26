import type { Document, HeaderFooter } from '../types/document';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import { parseDocumentWithRust } from './rustParseFacade';
import type { RustSaveDeterminism } from './rustSaveFacade';
import { isSessionRevisionInfo, withSessionRevisionIds } from './sessionRevisionIds';
import { unzipContainer } from './wasm';

const MODEL_PARTS_REWRITTEN = 'DOCX_MODEL_PARTS_REWRITTEN';
const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

function readSaveValue<T extends object, K extends keyof T>(object: T, key: K): T[K] {
  const descriptor = Object.getOwnPropertyDescriptor(object, key);
  if (descriptor ? !('value' in descriptor) : key in object) {
    throw new TypeError('Cannot compare an accessor or inherited save value');
  }
  return object[key];
}

function plainSaveMap<K, V>(entries: Map<K, V> | undefined): Map<K, V> | undefined {
  if (entries && (Object.getPrototypeOf(entries) !== Map.prototype || Reflect.ownKeys(entries).length)) {
    throw new TypeError('Cannot compare custom save maps');
  }
  return entries;
}

function headerFooterSaveEntries(
  entries: Map<string, HeaderFooter> | undefined
): Array<[string, HeaderFooter]> {
  return entries
    ? [...entries].filter(([, part]) => !isWrittenByCanonical(part, entries))
    : [];
}

function snapshot(value: unknown, active: WeakSet<object>): unknown {
  if (typeof value === 'function') throw new TypeError('Cannot compare a function save value');
  if (value === null || typeof value !== 'object') return value;
  if (active.has(value) || isSessionRevisionInfo(value) ||
      readSaveValue(value as { toJSON?: unknown }, 'toJSON') !== undefined) {
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
        { length: readSaveValue(value, 'length') },
        (_, index) => snapshot(readSaveValue(value, index), active)
      );
    }
    const record = value as Record<string, unknown>;
    const copy: Record<string, unknown> = Object.create(null);
    for (const key of Object.keys(record)) {
      if (FORBIDDEN_KEYS.has(key)) throw new TypeError('Cannot compare a forbidden save key');
      copy[key] = snapshot(readSaveValue(record, key), active);
    }
    return copy;
  } finally {
    active.delete(value);
  }
}

function saveModel(document: Document, determinism: RustSaveDeterminism): object {
  const pkg = readSaveValue(document, 'package');
  return {
    determinism,
    document: readSaveValue(pkg, 'document'),
    headerEntries: headerFooterSaveEntries(plainSaveMap(readSaveValue(pkg, 'headers'))),
    footerEntries: headerFooterSaveEntries(plainSaveMap(readSaveValue(pkg, 'footers'))),
    footnotes: readSaveValue(pkg, 'footnotes') ?? [],
    endnotes: readSaveValue(pkg, 'endnotes') ?? [],
    footnoteSeparators: readSaveValue(pkg, 'footnoteSeparators') ?? [],
    endnoteSeparators: readSaveValue(pkg, 'endnoteSeparators') ?? [],
    relationshipEntries: [...(plainSaveMap(readSaveValue(pkg, 'relationships'))?.entries() ?? [])],
    ...(readSaveValue(pkg, 'numbering') === undefined
      ? {}
      : { numbering: readSaveValue(pkg, 'numbering') }),
  };
}

export function modelMatchesSource(
  caller: Document,
  source: ArrayBuffer,
  determinism: RustSaveDeterminism
): boolean {
  try {
    const { document } = parseDocumentWithRust(source, {
      parseHeadersFooters: true,
      parseNotes: true,
      detectVariables: false,
      includeCanonical: false,
    });
    const original = saveModel(document, determinism);
    const originalWire = withSessionRevisionIds(original, () => JSON.stringify(original));
    const current = snapshot(saveModel(caller, determinism), new WeakSet()) as object;
    return withSessionRevisionIds(current, () => JSON.stringify(current)) === originalWire;
  } catch {
    return false;
  }
}

export function readModelSaveParts(buffer: ArrayBuffer): Map<string, Uint8Array> {
  const entries = unzipContainer(new Uint8Array(buffer));
  const parts = new Map(Object.entries(entries));
  // The ZIP reader's assignment to __proto__ can set the result object's prototype.
  const prototype = Object.getPrototypeOf(entries);
  if (prototype instanceof Uint8Array) parts.set('__proto__', prototype);
  return parts;
}

export function samePartBytes(before?: Uint8Array, after?: Uint8Array): boolean {
  if (before === after) return true;
  return before !== undefined && after !== undefined && before.byteLength === after.byteLength &&
    before.every((byte, index) => byte === after[index]);
}

export function modelPartsRewrittenWarning(
  original: Map<string, Uint8Array>,
  written: Map<string, Uint8Array>
): string | undefined {
  const paths = [...new Set([...original.keys(), ...written.keys()])]
    .filter((path) => !samePartBytes(original.get(path), written.get(path)))
    .sort();
  if (!paths.length) return undefined;
  return `${MODEL_PARTS_REWRITTEN} ${JSON.stringify(paths)} Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`;
}
