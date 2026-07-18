import type { Document, MediaFile } from '../../types/document';

const VERSION = 'docx-document-canonical-v1';
const UTF8 = new TextEncoder();
const DECODER = new TextDecoder();

type JsonObject = Record<string, unknown>;

export interface CanonicalComparison {
  equal: boolean;
  leftSha256: string;
  rightSha256: string;
  pointer?: string;
  leftSummary?: string;
  rightSummary?: string;
}

/** Project the incumbent parser's complete public Document without dropping fields. */
export function projectDocumentCanonical(document: Document): Document {
  assertMediaAliasIdentity(document.package.media);
  return document;
}

/** Encode the strict `docx-document-canonical-v1` byte contract. */
export function toDocumentCanonicalBytes(value: unknown): Uint8Array {
  const active = new WeakSet<object>();
  const json = encodeCanonicalValue(value, active, '');
  return UTF8.encode(`${VERSION}\n${json}\n`);
}

/** SHA-256 of exact bytes as lowercase hexadecimal. */
export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  // Copy onto a definite ArrayBuffer: the TS lib types conservatively allow a
  // Uint8Array to be backed by SharedArrayBuffer, which BufferSource rejects.
  const digest = await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/**
 * Byte-compare two canonical streams and, on mismatch, locate the first value
 * divergence as a JSON Pointer with compact type/value summaries.
 */
export async function compareDocumentCanonicalBytes(
  left: Uint8Array,
  right: Uint8Array
): Promise<CanonicalComparison> {
  const [leftSha256, rightSha256] = await Promise.all([sha256Hex(left), sha256Hex(right)]);
  if (bytesEqual(left, right)) return { equal: true, leftSha256, rightSha256 };

  const leftValue = parseCanonicalStream(left);
  const rightValue = parseCanonicalStream(right);
  const difference = firstDifference(leftValue, rightValue, '');
  return {
    equal: false,
    leftSha256,
    rightSha256,
    pointer: difference?.pointer ?? '',
    leftSummary: summarize(difference?.left),
    rightSummary: summarize(difference?.right),
  };
}

function encodeCanonicalValue(value: unknown, active: WeakSet<object>, pointer: string): string {
  if (value === null) return 'null';

  switch (typeof value) {
    case 'string':
    case 'boolean':
      return JSON.stringify(value);
    case 'number':
      if (!Number.isFinite(value)) contractError(pointer, 'non-finite number');
      return JSON.stringify(Object.is(value, -0) ? 0 : value);
    case 'undefined':
    case 'function':
    case 'symbol':
    case 'bigint':
      return contractError(pointer, typeof value);
    case 'object':
      break;
  }

  if (active.has(value)) contractError(pointer, 'cycle');
  active.add(value);
  try {
    if (value instanceof Date) {
      if (!Number.isFinite(value.getTime())) contractError(pointer, 'invalid Date');
      return `{\"$date\":${JSON.stringify(value.toISOString())}}`;
    }

    const bytes = asBytes(value);
    if (bytes) {
      const base64 = bytesToBase64(bytes);
      return `{\"$binary\":{\"base64\":${JSON.stringify(base64)},\"byteLength\":${bytes.byteLength}}}`;
    }

    if (value instanceof Map) {
      const seen = new Set<string>();
      const entries: string[] = [];
      let index = 0;
      for (const [key, entry] of value) {
        if (typeof key !== 'string')
          contractError(`${pointer}/$map/${index}/0`, 'non-string map key');
        if (seen.has(key)) contractError(`${pointer}/$map/${index}/0`, 'duplicate map key');
        seen.add(key);
        entries.push(
          `[${JSON.stringify(key)},${encodeCanonicalValue(entry, active, `${pointer}/$map/${index}/1`)}]`
        );
        index += 1;
      }
      return `{\"$map\":[${entries.join(',')}]}`;
    }

    if (Array.isArray(value)) {
      const entries: string[] = [];
      for (let index = 0; index < value.length; index += 1) {
        if (!Object.prototype.hasOwnProperty.call(value, index)) {
          contractError(`${pointer}/${index}`, 'array hole');
        }
        entries.push(encodeCanonicalValue(value[index], active, `${pointer}/${index}`));
      }
      return `[${entries.join(',')}]`;
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      contractError(pointer, `non-plain object ${prototype?.constructor?.name ?? '<null>'}`);
    }
    if (Object.getOwnPropertySymbols(value).length > 0)
      contractError(pointer, 'symbol-key property');

    const object = value as JsonObject;
    const keys = Object.keys(object)
      .filter((key) => object[key] !== undefined)
      .sort(compareUnicodeKeys);
    return `{${keys
      .map(
        (key) =>
          `${JSON.stringify(key)}:${encodeCanonicalValue(object[key], active, `${pointer}/${escapePointer(key)}`)}`
      )
      .join(',')}}`;
  } finally {
    active.delete(value);
  }
}

function asBytes(value: object): Uint8Array | null {
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  return null;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

function compareUnicodeKeys(left: string, right: string): number {
  const leftScalars = Array.from(left, (char) => char.codePointAt(0) ?? 0);
  const rightScalars = Array.from(right, (char) => char.codePointAt(0) ?? 0);
  const length = Math.min(leftScalars.length, rightScalars.length);
  for (let index = 0; index < length; index += 1) {
    const difference = leftScalars[index] - rightScalars[index];
    if (difference !== 0) return difference;
  }
  return leftScalars.length - rightScalars.length;
}

function contractError(pointer: string, kind: string): never {
  throw new TypeError(`canonical contract error at ${pointer || '/'}: ${kind}`);
}

function assertMediaAliasIdentity(media: Map<string, MediaFile> | undefined): void {
  if (!media) return;
  const byPath = new Map<string, MediaFile>();
  for (const [alias, file] of media) {
    const prior = byPath.get(file.path);
    if (prior && (prior !== file || prior.data !== file.data)) {
      throw new TypeError(`media alias ${JSON.stringify(alias)} does not reuse blob identity`);
    }
    byPath.set(file.path, file);
  }
}

function parseCanonicalStream(bytes: Uint8Array): unknown {
  const stream = DECODER.decode(bytes);
  const prefix = `${VERSION}\n`;
  if (!stream.startsWith(prefix) || !stream.endsWith('\n')) {
    throw new TypeError(`expected ${VERSION} canonical stream`);
  }
  return JSON.parse(stream.slice(prefix.length, -1)) as unknown;
}

function firstDifference(
  left: unknown,
  right: unknown,
  pointer: string
): { pointer: string; left: unknown; right: unknown } | null {
  if (Object.is(left, right)) return null;
  if (typeof left !== typeof right || left === null || right === null)
    return { pointer, left, right };

  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return { pointer, left, right };
    const length = Math.max(left.length, right.length);
    for (let index = 0; index < length; index += 1) {
      if (index >= left.length || index >= right.length) {
        return { pointer: `${pointer}/${index}`, left: left[index], right: right[index] };
      }
      const nested = firstDifference(left[index], right[index], `${pointer}/${index}`);
      if (nested) return nested;
    }
    return null;
  }

  if (typeof left === 'object' && typeof right === 'object') {
    const leftObject = left as JsonObject;
    const rightObject = right as JsonObject;
    const keys = Array.from(
      new Set([...Object.keys(leftObject), ...Object.keys(rightObject)])
    ).sort(compareUnicodeKeys);
    for (const key of keys) {
      if (!(key in leftObject) || !(key in rightObject)) {
        return {
          pointer: `${pointer}/${escapePointer(key)}`,
          left: leftObject[key],
          right: rightObject[key],
        };
      }
      const nested = firstDifference(
        leftObject[key],
        rightObject[key],
        `${pointer}/${escapePointer(key)}`
      );
      if (nested) return nested;
    }
    return null;
  }
  return { pointer, left, right };
}

function escapePointer(segment: string): string {
  return segment.replace(/~/g, '~0').replace(/\//g, '~1');
}

function summarize(value: unknown): string {
  if (value === undefined) return 'missing';
  if (value === null) return 'null';
  if (Array.isArray(value)) return `array(${value.length}) ${truncate(JSON.stringify(value))}`;
  if (typeof value === 'object') return `object ${truncate(JSON.stringify(value))}`;
  return `${typeof value} ${truncate(JSON.stringify(value))}`;
}

function truncate(value: string, length = 160): string {
  return value.length <= length ? value : `${value.slice(0, length - 1)}…`;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}
