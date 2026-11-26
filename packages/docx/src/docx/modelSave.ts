import type { Document, HeaderFooter } from '../types/document';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import { parseDocumentWithRust } from './rustParseFacade';
import type { RustSaveDeterminism } from './rustSaveFacade';
import { isSessionRevisionInfo, withSessionRevisionIds } from './sessionRevisionIds';
import { unzipContainer } from './wasm';

const MODEL_PARTS_REWRITTEN = 'DOCX_MODEL_PARTS_REWRITTEN';
const CORE_PROPERTIES = 'docProps/core.xml';
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

function zipEntryCount(buffer: ArrayBuffer): number | undefined {
  const view = new DataView(buffer);
  const footerSize = 22;
  const last = view.byteLength - footerSize;
  const first = Math.max(0, last - 0xffff);
  for (let at = last; at >= first; at--) {
    if (view.getUint32(at, true) !== 0x06054b50) continue;
    const end = at + footerSize + view.getUint16(at + 20, true);
    if (end > view.byteLength) continue;
    if (end < view.byteLength) return undefined;

    const count = view.getUint16(at + 10, true);
    const directorySize = view.getUint32(at + 12, true);
    const directoryOffset = view.getUint32(at + 16, true);
    if (
      view.getUint32(at + 4, true) !== 0 || view.getUint16(at + 8, true) !== count ||
      count === 0xffff || (at >= 20 && view.getUint32(at - 20, true) === 0x07064b50) ||
      directorySize < count * 46 || directoryOffset + directorySize !== at ||
      (count !== 0 && view.getUint32(directoryOffset, true) !== 0x02014b50)
    ) {
      return undefined;
    }
    return count;
  }
  return undefined;
}

export function readModelSaveParts(buffer: ArrayBuffer) {
  const entries = unzipContainer(new Uint8Array(buffer));
  const core = Object.hasOwn(entries, CORE_PROPERTIES) ? entries[CORE_PROPERTIES] : undefined;
  const members = Object.entries(entries);
  const reliable = Object.getPrototypeOf(entries) === Object.prototype &&
    members.length === zipEntryCount(buffer) &&
    members.every(([name]) => !(name in Object.prototype || name in Uint8Array.prototype));
  return { core, parts: reliable ? new Map(members) : undefined };
}

export function samePartBytes(before?: Uint8Array, after?: Uint8Array): boolean {
  if (before === after) return true;
  return before !== undefined && after !== undefined && before.byteLength === after.byteLength &&
    before.every((byte, index) => byte === after[index]);
}

export function modelPartsRewrittenWarning(
  original: Map<string, Uint8Array> | undefined,
  written: Map<string, Uint8Array> | undefined
): string | undefined {
  const guidance = 'Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.';
  if (!original || !written) return `${MODEL_PARTS_REWRITTEN} (part names unavailable) ${guidance}`;
  const paths = [...new Set([...original.keys(), ...written.keys()])]
    .filter((path) => !samePartBytes(original.get(path), written.get(path)))
    .sort();
  if (!paths.length) return undefined;
  return `${MODEL_PARTS_REWRITTEN} ${JSON.stringify(paths)} ${guidance}`;
}
