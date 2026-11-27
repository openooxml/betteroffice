import type { Document, HeaderFooter } from '../types/document';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import { parseDocumentWithRust } from './rustParseFacade';
import type { RustSaveDeterminism } from './rustSaveFacade';
import { isSessionRevisionInfo, withSessionRevisionIds } from './sessionRevisionIds';
import { unzipContainer } from './wasm';

const MODEL_PARTS_REWRITTEN = 'DOCX_MODEL_PARTS_REWRITTEN';
const CORE_PROPERTIES = 'docProps/core.xml';
const CONTENT_TYPES = '[Content_Types].xml';
const ZIP_FOOTER_SIZE = 22;
const ZIP_ENTRY_SIZE = 46;
const ZIP_LOCAL_HEADER_SIZE = 30;
const ZIP64_SENTINEL = 0xffffffff;
const ZIP_MEMBER_ALIAS = 'part';
const ZIP_MEMBER_ALIAS_BYTES = new TextEncoder().encode(ZIP_MEMBER_ALIAS);
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
    const current = saveModel(caller, determinism);
    const detached = snapshot(current, new WeakSet()) as object;
    return withSessionRevisionIds(detached, () => JSON.stringify(detached)) === originalWire &&
      withSessionRevisionIds(current, () => JSON.stringify(current)) === originalWire;
  } catch {
    return false;
  }
}

interface ZipMember {
  crc32: number;
  method: number;
  compressedSize: number;
  uncompressedSize: number;
  localOffset: number;
  centralOffset: number;
  duplicate?: boolean;
}

function zipDirectoryLocation(view: DataView, footer: number) {
  const disk = view.getUint16(footer + 4, true);
  const directoryDisk = view.getUint16(footer + 6, true);
  const locator = footer - 20;
  if (locator >= 0 && view.getUint32(locator, true) === 0x07064b50) {
    if (
      (disk !== 0 && disk !== 0xffff) || (directoryDisk !== 0 && directoryDisk !== 0xffff) ||
      view.getUint32(locator + 4, true) !== 0 || view.getUint32(locator + 16, true) !== 1
    ) return undefined;
    const record = Number(view.getBigUint64(locator + 8, true));
    if (
      record + 56 > locator ||
      view.getUint32(record, true) !== 0x06064b50 ||
      view.getBigUint64(record + 4, true) !== BigInt(locator - record - 12) ||
      view.getUint32(record + 16, true) !== 0 || view.getUint32(record + 20, true) !== 0 ||
      view.getBigUint64(record + 24, true) !== view.getBigUint64(record + 32, true)
    ) return undefined;
    return {
      count: Number(view.getBigUint64(record + 32, true)),
      size: Number(view.getBigUint64(record + 40, true)),
      offset: Number(view.getBigUint64(record + 48, true)),
      end: record,
    };
  }
  const count = view.getUint16(footer + 10, true);
  if (disk !== 0 || directoryDisk !== 0 || view.getUint16(footer + 8, true) !== count) {
    return undefined;
  }
  return {
    count,
    size: view.getUint32(footer + 12, true),
    offset: view.getUint32(footer + 16, true),
    end: footer,
  };
}

function findZipExtra(view: DataView, start: number, end: number, id: number) {
  for (let offset = start; offset + 4 <= end;) {
    const next = offset + 4 + view.getUint16(offset + 2, true);
    if (next > end) break;
    if (view.getUint16(offset, true) === id) return { start: offset + 4, end: next };
    offset = next;
  }
  return undefined;
}

function readZipMember(
  view: DataView,
  centralOffset: number,
  extraStart: number,
  extraEnd: number,
  archiveOffset: number
): ZipMember | undefined {
  let compressedSize = view.getUint32(centralOffset + 20, true);
  let uncompressedSize = view.getUint32(centralOffset + 24, true);
  let localOffset = view.getUint32(centralOffset + 42, true);
  let disk = view.getUint16(centralOffset + 34, true);
  if (
    compressedSize === ZIP64_SENTINEL || uncompressedSize === ZIP64_SENTINEL ||
    localOffset === ZIP64_SENTINEL || disk === 0xffff
  ) {
    const extra = findZipExtra(view, extraStart, extraEnd, 1);
    if (!extra) return undefined;
    let offset = extra.start;
    const takeZip64 = (value: number) => {
      if (value !== ZIP64_SENTINEL) return value;
      if (offset + 8 > extra.end) return NaN;
      const size = Number(view.getBigUint64(offset, true));
      offset += 8;
      return size;
    };
    uncompressedSize = takeZip64(uncompressedSize);
    compressedSize = takeZip64(compressedSize);
    localOffset = takeZip64(localOffset);
    if (disk === 0xffff && offset + 4 <= extra.end) disk = view.getUint32(offset, true);
  }
  if (
    disk !== 0 || !Number.isFinite(uncompressedSize) ||
    !(localOffset + archiveOffset + ZIP_LOCAL_HEADER_SIZE + compressedSize <= centralOffset)
  ) return undefined;
  return {
    crc32: view.getUint32(centralOffset + 16, true),
    method: view.getUint16(centralOffset + 10, true),
    compressedSize,
    uncompressedSize,
    localOffset: localOffset + archiveOffset,
    centralOffset,
  };
}

function filenameCrc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
  }
  return ~crc >>> 0;
}

function readZipName(
  view: DataView<ArrayBuffer>,
  decoder: TextDecoder,
  offset: number,
  extraStart: number,
  extraEnd: number
): string | undefined {
  const nameOffset = offset + ZIP_ENTRY_SIZE;
  const bytes = new Uint8Array(view.buffer, nameOffset, extraStart - nameOffset);
  const unicode = findZipExtra(view, extraStart, extraEnd, 0x7075);
  const utf8 = (view.getUint16(offset + 8, true) & 0x0800) !== 0;
  try {
    if (
      unicode && unicode.end - unicode.start > 5 && view.getUint8(unicode.start) === 1 &&
      view.getUint32(unicode.start + 1, true) === filenameCrc32(bytes)
    ) {
      return decoder.decode(new Uint8Array(view.buffer, unicode.start + 5, unicode.end - unicode.start - 5));
    }
    if (utf8) return decoder.decode(bytes);
  } catch {
    return undefined;
  }
  return Array.from(bytes, (byte) =>
    byte < 0x80 ? String.fromCharCode(byte) : CP437_HIGH_BYTES[byte - 0x80]
  ).join('');
}

export function readModelSaveDirectory(buffer: ArrayBuffer): Map<string, ZipMember> | undefined {
  const view = new DataView(buffer);
  const last = view.byteLength - ZIP_FOOTER_SIZE;
  const first = Math.max(0, last - 0xffff);
  let footer = last;
  for (; footer >= first; footer--) {
    if (
      view.getUint32(footer, true) === 0x06054b50 &&
      footer + ZIP_FOOTER_SIZE + view.getUint16(footer + 20, true) === view.byteLength
    ) {
      break;
    }
  }
  if (footer < first) return undefined;

  const directory = zipDirectoryLocation(view, footer);
  if (!directory || directory.count * ZIP_ENTRY_SIZE > directory.size) return undefined;
  let offset = directory.end - directory.size;
  const archiveOffset = offset - directory.offset;
  if (archiveOffset < 0) return undefined;

  const members = new Map<string, ZipMember>();
  const decoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
  for (let entry = 0; entry < directory.count; entry++) {
    if (offset + ZIP_ENTRY_SIZE > directory.end || view.getUint32(offset, true) !== 0x02014b50) {
      return undefined;
    }
    const nameSize = view.getUint16(offset + 28, true);
    const extraStart = offset + ZIP_ENTRY_SIZE + nameSize;
    const extraEnd = extraStart + view.getUint16(offset + 30, true);
    const next = extraEnd + view.getUint16(offset + 32, true);
    if (nameSize === 0 || next > directory.end) return undefined;
    const member = readZipMember(view, offset, extraStart, extraEnd, archiveOffset);
    if (!member) return undefined;

    const name = readZipName(view, decoder, offset, extraStart, extraEnd);
    if (name === undefined) return undefined;
    if (members.has(name)) member.duplicate = true;
    members.set(name, member);
    offset = next;
  }
  if (offset === directory.end) return members;
  if (
    offset + 6 <= directory.end && view.getUint32(offset, true) === 0x05054b50 &&
    offset + 6 + view.getUint16(offset + 4, true) === directory.end
  ) return members;
  return undefined;
}

function sameBytes(before: Uint8Array, after: Uint8Array): boolean {
  return before.byteLength === after.byteLength && before.every((byte, index) => byte === after[index]);
}

function readZipPayload(view: DataView<ArrayBuffer>, member: ZipMember): Uint8Array | undefined {
  const local = member.localOffset;
  if (
    view.getUint32(local, true) !== 0x04034b50 ||
    view.getUint16(local + 8, true) !== member.method ||
    ((view.getUint16(local + 6, true) | view.getUint16(member.centralOffset + 8, true)) & 1)
  ) return undefined;
  const dataStart = local + ZIP_LOCAL_HEADER_SIZE +
    view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
  if (dataStart + member.compressedSize > member.centralOffset) return undefined;
  return new Uint8Array(view.buffer, dataStart, member.compressedSize);
}

function unpackZipPayload(member: ZipMember, payload: Uint8Array): Uint8Array | undefined {
  if (member.compressedSize >= ZIP64_SENTINEL || member.uncompressedSize >= ZIP64_SENTINEL) {
    return undefined;
  }

  // Decode a single payload under a fixed name because the WASM reader returns a plain object.
  const nameSize = ZIP_MEMBER_ALIAS_BYTES.byteLength;
  const dataStart = ZIP_LOCAL_HEADER_SIZE + nameSize;
  const central = dataStart + payload.byteLength;
  const centralSize = ZIP_ENTRY_SIZE + nameSize;
  const footer = central + centralSize;
  const single = new Uint8Array(footer + ZIP_FOOTER_SIZE);
  single.set(ZIP_MEMBER_ALIAS_BYTES, ZIP_LOCAL_HEADER_SIZE);
  single.set(payload, dataStart);
  single.set(ZIP_MEMBER_ALIAS_BYTES, central + ZIP_ENTRY_SIZE);
  const output = new DataView(single.buffer);
  output.setUint32(0, 0x04034b50, true);
  output.setUint16(4, 20, true);
  output.setUint16(8, member.method, true);
  output.setUint32(14, member.crc32, true);
  output.setUint32(18, member.compressedSize, true);
  output.setUint32(22, member.uncompressedSize, true);
  output.setUint16(26, nameSize, true);
  output.setUint32(central, 0x02014b50, true);
  output.setUint16(central + 4, 20, true);
  output.setUint16(central + 6, 20, true);
  output.setUint16(central + 10, member.method, true);
  output.setUint32(central + 16, member.crc32, true);
  output.setUint32(central + 20, member.compressedSize, true);
  output.setUint32(central + 24, member.uncompressedSize, true);
  output.setUint16(central + 28, nameSize, true);
  output.setUint32(footer, 0x06054b50, true);
  output.setUint16(footer + 8, 1, true);
  output.setUint16(footer + 10, 1, true);
  output.setUint32(footer + 12, centralSize, true);
  output.setUint32(footer + 16, central, true);

  const parts = unzipContainer(single);
  const bytes = Object.prototype.hasOwnProperty.call(parts, ZIP_MEMBER_ALIAS)
    ? parts[ZIP_MEMBER_ALIAS]
    : undefined;
  return bytes?.byteLength === member.uncompressedSize ? bytes : undefined;
}

export function sameModelSavePackageMetadata(original: ArrayBuffer, written: ArrayBuffer): boolean {
  try {
    const before = readModelSaveDirectory(original);
    const after = readModelSaveDirectory(written);
    if (!before || !after) return false;
    const originalView = new DataView(original);
    const writtenView = new DataView(written);
    for (const path of new Set([...before.keys(), ...after.keys()])) {
      if (path !== CORE_PROPERTIES && path !== CONTENT_TYPES && !path.endsWith('.rels')) continue;
      const old = before.get(path);
      const current = after.get(path);
      if (!old || !current || old.duplicate || current.duplicate) return false;
      const oldPayload = readZipPayload(originalView, old);
      const currentPayload = readZipPayload(writtenView, current);
      const oldBytes = oldPayload && unpackZipPayload(old, oldPayload);
      const currentBytes = currentPayload && unpackZipPayload(current, currentPayload);
      if (!oldBytes || !currentBytes || !sameBytes(oldBytes, currentBytes)) return false;
    }
    return true;
  } catch {
    return false;
  }
}

export function modelPartsRewrittenWarning(
  original: ArrayBuffer,
  written: ArrayBuffer
): string | undefined {
  if (sameBytes(new Uint8Array(original), new Uint8Array(written))) return undefined;
  const guidance = 'Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.';
  const unavailable = `${MODEL_PARTS_REWRITTEN} (part names unavailable) ${guidance}`;
  try {
    const before = readModelSaveDirectory(original);
    const after = readModelSaveDirectory(written);
    if (!before || !after) return unavailable;
    const originalView = new DataView(original);
    const writtenView = new DataView(written);
    const paths = [...new Set([...before.keys(), ...after.keys()])]
      .filter((path) => {
        const old = before.get(path);
        const current = after.get(path);
        if (
          !old || !current || old.duplicate || current.duplicate ||
          old.crc32 !== current.crc32 || old.uncompressedSize !== current.uncompressedSize
        ) return true;
        const oldPayload = readZipPayload(originalView, old);
        const currentPayload = readZipPayload(writtenView, current);
        if (!oldPayload || !currentPayload) return true;
        if (old.method === current.method && sameBytes(oldPayload, currentPayload)) return false;
        const oldBytes = unpackZipPayload(old, oldPayload);
        const currentBytes = unpackZipPayload(current, currentPayload);
        return !oldBytes || !currentBytes || !sameBytes(oldBytes, currentBytes);
      })
      .sort();
    return paths.length ? `${MODEL_PARTS_REWRITTEN} ${JSON.stringify(paths)} ${guidance}` : undefined;
  } catch {
    return unavailable;
  }
}
