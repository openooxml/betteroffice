import type { Document, HeaderFooter } from '../types/document';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import { parseDocumentWithRust } from './rustParseFacade';
import type { RustSaveDeterminism } from './rustSaveFacade';
import { isSessionRevisionInfo, withSessionRevisionIds } from './sessionRevisionIds';
import { unzipContainer } from './wasm';

const MODEL_PARTS_REWRITTEN = 'DOCX_MODEL_PARTS_REWRITTEN';
const CORE_PROPERTIES = 'docProps/core.xml';
const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
const CP437_HIGH_BYTES = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»' +
  '░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀' +
  'αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\u00a0';

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

interface ZipMember {
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  method: number;
}

export function readModelSaveDirectory(buffer: ArrayBuffer): Map<string, ZipMember> | undefined {
  const view = new DataView(buffer);
  const footerSize = 22;
  const last = view.byteLength - footerSize;
  const first = Math.max(0, last - 0xffff);
  let footer = last;
  for (; footer >= first; footer--) {
    if (
      view.getUint32(footer, true) === 0x06054b50 &&
      footer + footerSize + view.getUint16(footer + 20, true) === view.byteLength
    ) {
      break;
    }
  }
  if (footer < first) return undefined;

  const entrySize = 46;
  const count = view.getUint16(footer + 10, true);
  const directorySize = view.getUint32(footer + 12, true);
  let offset = view.getUint32(footer + 16, true);
  if (
    view.getUint32(footer + 4, true) !== 0 || view.getUint16(footer + 8, true) !== count ||
    count === 0xffff || directorySize === 0xffffffff || offset === 0xffffffff ||
    (footer >= 20 && view.getUint32(footer - 20, true) === 0x07064b50) ||
    directorySize < count * entrySize || offset + directorySize !== footer
  ) {
    return undefined;
  }

  const members = new Map<string, ZipMember>();
  const decoder = new TextDecoder('utf-8', { ignoreBOM: true });
  for (let entry = 0; entry < count; entry++) {
    if (offset + entrySize > footer || view.getUint32(offset, true) !== 0x02014b50) return undefined;
    const nameSize = view.getUint16(offset + 28, true);
    const next = offset + entrySize + nameSize +
      view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    if (
      nameSize === 0 || next > footer || view.getUint16(offset + 34, true) !== 0 ||
      view.getUint32(offset + 42, true) === 0xffffffff ||
      compressedSize === 0xffffffff || uncompressedSize === 0xffffffff
    ) {
      return undefined;
    }

    const nameBytes = new Uint8Array(buffer, offset + entrySize, nameSize);
    const utf8 = (view.getUint16(offset + 8, true) & 0x0800) !== 0;
    const name = utf8
      ? decoder.decode(nameBytes)
      : Array.from(nameBytes, (byte) => byte < 0x80 ? String.fromCharCode(byte) : CP437_HIGH_BYTES[byte - 0x80]).join('');
    if (utf8 && name.includes('\ufffd')) {
      const encoded = new TextEncoder().encode(name);
      if (encoded.length !== nameBytes.length || !nameBytes.every((byte, index) => byte === encoded[index])) {
        return undefined;
      }
    }
    if (members.has(name)) return undefined;
    members.set(name, {
      crc32: view.getUint32(offset + 16, true),
      compressedSize,
      uncompressedSize,
      method: view.getUint16(offset + 10, true),
    });
    offset = next;
  }
  return offset === footer ? members : undefined;
}

function readModelSaveCore(buffer: ArrayBuffer): Uint8Array | undefined {
  const entries = unzipContainer(new Uint8Array(buffer));
  return Object.hasOwn(entries, CORE_PROPERTIES) ? entries[CORE_PROPERTIES] : undefined;
}

export function sameModelSaveCore(original: ArrayBuffer, written: ArrayBuffer): boolean {
  const before = readModelSaveCore(original);
  const after = readModelSaveCore(written);
  if (before === after) return true;
  return before !== undefined && after !== undefined && before.byteLength === after.byteLength &&
    before.every((byte, index) => byte === after[index]);
}

export function modelPartsRewrittenWarning(
  original: ArrayBuffer,
  written: ArrayBuffer
): string | undefined {
  const guidance = 'Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.';
  const before = readModelSaveDirectory(original);
  const after = readModelSaveDirectory(written);
  if (!before || !after) return `${MODEL_PARTS_REWRITTEN} (part names unavailable) ${guidance}`;
  const paths = [...new Set([...before.keys(), ...after.keys()])]
    .filter((path) => {
      const old = before.get(path);
      const current = after.get(path);
      return !old || !current || old.crc32 !== current.crc32 || old.uncompressedSize !== current.uncompressedSize;
    })
    .sort();
  if (!paths.length) return undefined;
  return `${MODEL_PARTS_REWRITTEN} ${JSON.stringify(paths)} ${guidance}`;
}
