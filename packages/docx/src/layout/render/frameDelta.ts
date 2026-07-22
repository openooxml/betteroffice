import type { DisplayList, DisplayPage, DisplayPrimitive } from './displayList';

export const FRAME_DELTA_VERSION = 1;
export const FRAME_DELTA_HEADER_BYTES = 80;
export const FRAME_DELTA_PAGE_OP_BYTES = 48;

const FRAME_FLAG_FULL = 1;
const PAGE_OP_UPSERT = 1;
const PAGE_OP_REMOVE = 2;
const PAGE_OP_MOVE = 3;
const PAGE_OP_PATCH_POSITIONS = 4;
const PAGE_OP_SHIFT_POSITIONS = 5;
const POSITION_DOC_START = 1 << 0;
const POSITION_DOC_END = 1 << 1;
const POSITION_FRAGMENT_START = 1 << 2;
const POSITION_FRAGMENT_END = 1 << 3;
const POSITION_INLINE_WIDGET = 1 << 4;
const POSITION_MASK =
  POSITION_DOC_START |
  POSITION_DOC_END |
  POSITION_FRAGMENT_START |
  POSITION_FRAGMENT_END |
  POSITION_INLINE_WIDGET;
const MAX_VALUE_DEPTH = 64;
const MAX_CONTAINER_ITEMS = 10_000_000;
const MAX_SAFE_BIGINT = BigInt(Number.MAX_SAFE_INTEGER);
const MIN_SAFE_BIGINT = BigInt(Number.MIN_SAFE_INTEGER);

const VALUE_NULL = 0;
const VALUE_FALSE = 1;
const VALUE_TRUE = 2;
const VALUE_I64 = 3;
const VALUE_U64 = 4;
const VALUE_F64 = 5;
const VALUE_STRING = 6;
const VALUE_ARRAY = 7;
const VALUE_OBJECT = 8;
const VALUE_GLYPH_ARRAY = 9;
const GLYPH_LOGICAL_ORDER = 1 << 0;
const GLYPH_BIDI_LEVEL = 1 << 1;

export type FrameDeltaErrorCode =
  | 'invalid-frame'
  | 'unsupported-version'
  | 'stale-frame'
  | 'base-mismatch';

export class FrameDeltaError extends Error {
  readonly code: FrameDeltaErrorCode;

  constructor(code: FrameDeltaErrorCode, message: string) {
    super(`FrameDelta: ${message}`);
    this.name = 'FrameDeltaError';
    this.code = code;
  }
}

export interface FramePageUpsert {
  readonly kind: 'upsert';
  readonly pageIndex: number;
  readonly pageId: bigint;
  readonly fingerprint: bigint;
  /** Zero-copy view into {@link DecodedFrameDelta.bytes}. */
  readonly primitiveIds: BigUint64Array;
  readonly page: DisplayPage;
}

export interface FramePageRemove {
  readonly kind: 'remove';
  readonly pageIndex: number;
  readonly pageId: bigint;
}

export interface FramePageMove {
  readonly kind: 'move';
  readonly pageIndex: number;
  readonly pageId: bigint;
  readonly fingerprint: bigint;
}

export interface FramePrimitivePositionPatch {
  readonly primitiveId: bigint;
  readonly docStart?: number | null;
  readonly docEnd?: number | null;
  readonly fragmentDocStart?: number | null;
  readonly fragmentDocEnd?: number | null;
  readonly inlineWidgetPos?: number | null;
}

export interface FramePagePositionPatch {
  readonly kind: 'patch-positions';
  readonly pageIndex: number;
  readonly pageId: bigint;
  readonly fingerprint: bigint;
  readonly patches: readonly FramePrimitivePositionPatch[];
}

export interface FramePositionShiftRun {
  readonly start: number;
  readonly count: number;
  readonly changedMask: number;
  readonly delta: number;
}

export interface FramePagePositionShift {
  readonly kind: 'shift-positions';
  readonly pageIndex: number;
  readonly pageId: bigint;
  readonly fingerprint: bigint;
  readonly runs: readonly FramePositionShiftRun[];
}

export type FramePageOperation =
  | FramePageUpsert
  | FramePageRemove
  | FramePageMove
  | FramePagePositionPatch
  | FramePagePositionShift;

export interface DecodedFrameDelta {
  readonly protocolVersion: typeof FRAME_DELTA_VERSION;
  readonly full: boolean;
  readonly docEpoch: number;
  readonly layoutEpoch: number;
  readonly frameEpoch: number;
  readonly baseFrameEpoch: number;
  readonly pageCount: number;
  readonly contractVersion?: number;
  readonly operations: readonly FramePageOperation[];
  /** Complete binary frame; safe to transfer when the caller owns its buffer. */
  readonly bytes: Uint8Array;
}

export interface RetainedFramePage {
  readonly pageIndex: number;
  readonly pageId: bigint;
  readonly fingerprint: bigint;
  readonly primitiveIds: BigUint64Array;
  readonly page: DisplayPage;
}

export interface RetainedFrame {
  readonly protocolVersion: typeof FRAME_DELTA_VERSION;
  readonly docEpoch: number;
  readonly layoutEpoch: number;
  readonly frameEpoch: number;
  readonly contractVersion?: number;
  readonly pages: readonly RetainedFramePage[];
  /** IDs whose canvas/mirror/interactive projections must be replaced. */
  readonly damagedPageIds: ReadonlySet<bigint>;
  /** Removed surfaces, retained separately because they are absent from pages. */
  readonly removedPageIds: ReadonlySet<bigint>;
  readonly displayList: DisplayList;
}

interface RawPageOp {
  opcode: number;
  pageIndex: number;
  pageId: bigint;
  fingerprint: bigint;
  primitiveCount: number;
  primitiveIdsOffset: number;
  payloadOffset: number;
  payloadLength: number;
}

/**
 * Decode and fully validate one binary FrameDelta v1 without mutating retained
 * browser state. Page payloads are typed values, not embedded JSON strings.
 */
export function decodeFrameDelta(input: Uint8Array | ArrayBuffer): DecodedFrameDelta {
  let bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  // Primitive ids are intentionally 8-byte aligned relative to frame start.
  // wasm-bindgen returns offset-zero Uint8Arrays; normalize unusual subarrays
  // once so every exposed BigUint64Array remains a zero-copy view thereafter.
  if (bytes.byteOffset % 8 !== 0) bytes = bytes.slice();
  const reader = new BinaryReader(bytes);
  reader.requireLength(FRAME_DELTA_HEADER_BYTES, 'fixed header');
  if (
    reader.u8(0) !== 0x46 ||
    reader.u8(1) !== 0x44 ||
    reader.u8(2) !== 0x56 ||
    reader.u8(3) !== 0x31
  ) {
    invalid('bad magic');
  }
  const protocolVersion = reader.u16(4);
  if (protocolVersion !== FRAME_DELTA_VERSION) {
    throw new FrameDeltaError(
      'unsupported-version',
      `unsupported protocol version ${protocolVersion}`
    );
  }
  if (reader.u16(6) !== FRAME_DELTA_HEADER_BYTES) invalid('invalid v1 header length');
  if (reader.u32(8) !== bytes.byteLength) invalid('declared total length does not match buffer');
  const flags = reader.u32(12);
  if ((flags & ~FRAME_FLAG_FULL) !== 0) invalid(`unknown frame flags 0x${flags.toString(16)}`);
  const full = (flags & FRAME_FLAG_FULL) !== 0;
  const docEpoch = reader.safeU64(16, 'document epoch');
  const layoutEpoch = reader.safeU64(24, 'layout epoch');
  const frameEpoch = reader.safeU64(32, 'frame epoch');
  const baseFrameEpoch = reader.safeU64(40, 'base frame epoch');
  if (full && baseFrameEpoch !== 0) invalid('full frame must have base epoch zero');
  if (!full && baseFrameEpoch === 0) invalid('delta frame requires a nonzero base epoch');
  const pageCount = reader.u32(48);
  const operationCount = reader.u32(52);
  const operationsOffset = reader.u32(56);
  const stringsOffset = reader.u32(60);
  const stringsLength = reader.u32(64);
  const dataOffset = reader.u32(68);
  const encodedContractVersion = reader.u32(72);
  const contractVersion = encodedContractVersion === 0 ? undefined : encodedContractVersion;
  if (reader.u32(76) !== 0) invalid('reserved header bytes must be zero');
  if (operationsOffset !== FRAME_DELTA_HEADER_BYTES) invalid('invalid operation table offset');
  const operationsEnd = checkedAdd(
    operationsOffset,
    checkedMultiply(operationCount, FRAME_DELTA_PAGE_OP_BYTES, 'operation table'),
    'operation table'
  );
  if (operationsEnd !== stringsOffset) invalid('operation table/string table are not contiguous');
  const stringsEnd = checkedAdd(stringsOffset, stringsLength, 'string table');
  if (stringsEnd > dataOffset || dataOffset > bytes.byteLength) {
    invalid('string/data section bounds are invalid');
  }
  if (dataOffset % 8 !== 0) invalid('data section is not 8-byte aligned');

  const strings = decodeStringTable(reader, stringsOffset, stringsEnd);
  const rawOperations: RawPageOp[] = [];
  const pageIds = new Set<bigint>();
  for (let index = 0; index < operationCount; index++) {
    const offset = operationsOffset + index * FRAME_DELTA_PAGE_OP_BYTES;
    const opcode = reader.u8(offset);
    if (
      opcode !== PAGE_OP_UPSERT &&
      opcode !== PAGE_OP_REMOVE &&
      opcode !== PAGE_OP_MOVE &&
      opcode !== PAGE_OP_PATCH_POSITIONS &&
      opcode !== PAGE_OP_SHIFT_POSITIONS
    ) {
      invalid(`unknown page opcode ${opcode}`);
    }
    for (let reserved = 1; reserved < 4; reserved++) {
      if (reader.u8(offset + reserved) !== 0) invalid('page operation reserved bytes are nonzero');
    }
    for (let reserved = 40; reserved < FRAME_DELTA_PAGE_OP_BYTES; reserved++) {
      if (reader.u8(offset + reserved) !== 0) invalid('page operation tail is nonzero');
    }
    const pageIndex = reader.u32(offset + 4);
    const pageId = reader.u64(offset + 8);
    const fingerprint = reader.u64(offset + 16);
    const primitiveCount = reader.u32(offset + 24);
    const primitiveIdsOffset = reader.u32(offset + 28);
    const payloadOffset = reader.u32(offset + 32);
    const payloadLength = reader.u32(offset + 36);
    if (pageId === 0n) invalid('page id zero is reserved');
    if (pageIds.has(pageId)) invalid(`duplicate page operation for id ${pageId}`);
    pageIds.add(pageId);
    if (opcode === PAGE_OP_REMOVE || opcode === PAGE_OP_MOVE) {
      if (primitiveCount || primitiveIdsOffset || payloadOffset || payloadLength) {
        invalid('remove/move operation carries an unexpected payload');
      }
    } else if (opcode === PAGE_OP_UPSERT) {
      if (pageIndex >= pageCount) invalid('upsert page index exceeds final page count');
      if (primitiveIdsOffset < dataOffset || primitiveIdsOffset % 8 !== 0) {
        invalid('primitive id array has an invalid offset/alignment');
      }
      const primitiveIdsEnd = checkedAdd(
        primitiveIdsOffset,
        checkedMultiply(primitiveCount, 8, 'primitive id array'),
        'primitive id array'
      );
      const payloadEnd = checkedAdd(payloadOffset, payloadLength, 'page payload');
      if (
        primitiveIdsEnd > bytes.byteLength ||
        payloadOffset < primitiveIdsEnd ||
        payloadEnd > bytes.byteLength ||
        payloadLength === 0
      ) {
        invalid('page payload bounds are invalid');
      }
    } else {
      if (pageIndex >= pageCount) invalid('position patch page index exceeds final page count');
      const payloadEnd = checkedAdd(payloadOffset, payloadLength, 'position patch payload');
      if (
        primitiveCount > MAX_CONTAINER_ITEMS ||
        primitiveIdsOffset !== 0 ||
        payloadOffset < dataOffset ||
        payloadOffset % 8 !== 0 ||
        payloadLength < 8 ||
        payloadEnd > bytes.byteLength
      ) {
        invalid('position patch payload bounds are invalid');
      }
    }
    rawOperations.push({
      opcode,
      pageIndex,
      pageId,
      fingerprint,
      primitiveCount,
      primitiveIdsOffset,
      payloadOffset,
      payloadLength,
    });
  }
  if (full && rawOperations.some((operation) => operation.opcode !== PAGE_OP_UPSERT)) {
    invalid('full frame may contain only page upserts');
  }
  validateDataRegions(reader, rawOperations, stringsEnd, dataOffset);

  const operations = rawOperations.map((operation): FramePageOperation => {
    if (operation.opcode === PAGE_OP_REMOVE) {
      return { kind: 'remove', pageIndex: operation.pageIndex, pageId: operation.pageId };
    }
    if (operation.opcode === PAGE_OP_MOVE) {
      if (operation.pageIndex >= pageCount) invalid('move page index exceeds final page count');
      return {
        kind: 'move',
        pageIndex: operation.pageIndex,
        pageId: operation.pageId,
        fingerprint: operation.fingerprint,
      };
    }
    if (operation.opcode === PAGE_OP_PATCH_POSITIONS) {
      return {
        kind: 'patch-positions',
        pageIndex: operation.pageIndex,
        pageId: operation.pageId,
        fingerprint: operation.fingerprint,
        patches: decodePositionPatches(reader, operation),
      };
    }
    if (operation.opcode === PAGE_OP_SHIFT_POSITIONS) {
      return {
        kind: 'shift-positions',
        pageIndex: operation.pageIndex,
        pageId: operation.pageId,
        fingerprint: operation.fingerprint,
        runs: decodePositionShiftRuns(reader, operation),
      };
    }
    const payloadEnd = operation.payloadOffset + operation.payloadLength;
    const cursor = new ValueCursor(reader, strings, operation.payloadOffset, payloadEnd);
    const page = cursor.value(0) as DisplayPage;
    if (cursor.offset !== payloadEnd) invalid('page payload has trailing bytes');
    validateDisplayPage(page, operation.pageIndex, operation.primitiveCount);
    const absoluteIdsOffset = bytes.byteOffset + operation.primitiveIdsOffset;
    if (absoluteIdsOffset % 8 !== 0) invalid('primitive id array is not host-view aligned');
    const primitiveIds = new BigUint64Array(
      bytes.buffer,
      absoluteIdsOffset,
      operation.primitiveCount
    );
    const seenPrimitiveIds = new Set<bigint>();
    for (const primitiveId of primitiveIds) {
      if (primitiveId === 0n || seenPrimitiveIds.has(primitiveId)) {
        invalid('primitive id array contains zero or duplicate ids');
      }
      seenPrimitiveIds.add(primitiveId);
    }
    return {
      kind: 'upsert',
      pageIndex: operation.pageIndex,
      pageId: operation.pageId,
      fingerprint: operation.fingerprint,
      primitiveIds,
      page,
    };
  });
  if (full && operations.length !== pageCount) invalid('full frame does not define every page');

  return {
    protocolVersion: FRAME_DELTA_VERSION,
    full,
    docEpoch,
    layoutEpoch,
    frameEpoch,
    baseFrameEpoch,
    pageCount,
    contractVersion,
    operations,
    bytes,
  };
}

/** Apply one already-validated frame atomically, rejecting stale generations. */
export function applyFrameDelta(
  previous: RetainedFrame | null,
  delta: DecodedFrameDelta
): RetainedFrame {
  return applyFrameDeltaInternal(previous, delta, false);
}

/**
 * Apply a frame while reusing position-only page arenas owned exclusively by
 * the caller. Worker/main resident-engine hot paths replace their outer frame
 * atomically and have no concurrent readers; avoiding deep suffix clones keeps
 * geometry shifts from creating periodic multi-megabyte GC stalls.
 */
export function applyFrameDeltaOwned(
  previous: RetainedFrame | null,
  delta: DecodedFrameDelta
): RetainedFrame {
  return applyFrameDeltaInternal(previous, delta, true);
}

function applyFrameDeltaInternal(
  previous: RetainedFrame | null,
  delta: DecodedFrameDelta,
  reusePositionPages: boolean
): RetainedFrame {
  if (previous && delta.frameEpoch <= previous.frameEpoch) {
    throw new FrameDeltaError(
      'stale-frame',
      `frame ${delta.frameEpoch} is not newer than applied frame ${previous.frameEpoch}`
    );
  }
  if (!delta.full) {
    if (!previous || delta.baseFrameEpoch !== previous.frameEpoch) {
      throw new FrameDeltaError(
        'base-mismatch',
        `delta base ${delta.baseFrameEpoch} does not match applied frame ${previous?.frameEpoch ?? 0}`
      );
    }
  }
  if (previous && delta.docEpoch < previous.docEpoch) {
    throw new FrameDeltaError('stale-frame', 'document epoch moved backwards');
  }
  if (previous && delta.layoutEpoch < previous.layoutEpoch) {
    throw new FrameDeltaError('stale-frame', 'layout epoch moved backwards');
  }

  const pages = new Map<bigint, RetainedFramePage>();
  if (!delta.full && previous) {
    for (const page of previous.pages) pages.set(page.pageId, page);
  }
  const damagedPageIds = new Set<bigint>();
  const removedPageIds = new Set<bigint>();
  for (const operation of delta.operations) {
    if (operation.kind === 'remove') {
      if (!pages.delete(operation.pageId))
        invalid(`remove references unknown page ${operation.pageId}`);
      removedPageIds.add(operation.pageId);
      continue;
    }
    if (operation.kind === 'move') {
      const current = pages.get(operation.pageId);
      if (!current) invalid(`move references unknown page ${operation.pageId}`);
      if (current.fingerprint !== operation.fingerprint) {
        invalid(`move fingerprint differs for page ${operation.pageId}`);
      }
      pages.set(operation.pageId, {
        ...current,
        pageIndex: operation.pageIndex,
        page:
          current.page.pageIndex === operation.pageIndex
            ? current.page
            : { ...current.page, pageIndex: operation.pageIndex },
      });
      continue;
    }
    if (operation.kind === 'patch-positions') {
      const current = pages.get(operation.pageId);
      if (!current) invalid(`position patch references unknown page ${operation.pageId}`);
      pages.set(operation.pageId, {
        ...current,
        pageIndex: operation.pageIndex,
        fingerprint: operation.fingerprint,
        page: patchDisplayPagePositions(
          current.page,
          current.primitiveIds,
          operation.patches,
          operation.pageIndex
        ),
      });
      continue;
    }
    if (operation.kind === 'shift-positions') {
      const current = pages.get(operation.pageId);
      if (!current) invalid(`position shift references unknown page ${operation.pageId}`);
      pages.set(operation.pageId, {
        ...current,
        pageIndex: operation.pageIndex,
        fingerprint: operation.fingerprint,
        page: reusePositionPages
          ? shiftDisplayPagePositionsOwned(current.page, operation.runs, operation.pageIndex)
          : shiftDisplayPagePositions(current.page, operation.runs, operation.pageIndex),
      });
      continue;
    }
    pages.set(operation.pageId, {
      pageIndex: operation.pageIndex,
      pageId: operation.pageId,
      fingerprint: operation.fingerprint,
      primitiveIds: operation.primitiveIds,
      page: operation.page,
    });
    damagedPageIds.add(operation.pageId);
  }
  if (delta.full && previous) {
    for (const previousPage of previous.pages) {
      if (!pages.has(previousPage.pageId)) removedPageIds.add(previousPage.pageId);
    }
  }
  if (pages.size !== delta.pageCount) invalid('applied page count does not match frame header');
  const ordered = [...pages.values()].sort((left, right) => left.pageIndex - right.pageIndex);
  const orderedIds = new Set<bigint>();
  for (let index = 0; index < ordered.length; index++) {
    const page = ordered[index];
    if (page.pageIndex !== index || page.page.pageIndex !== index) {
      invalid('applied pages are not a contiguous zero-based sequence');
    }
    if (orderedIds.has(page.pageId)) invalid(`duplicate retained page id ${page.pageId}`);
    orderedIds.add(page.pageId);
  }
  const displayList: DisplayList = {
    ...(delta.contractVersion === undefined ? {} : { contractVersion: delta.contractVersion }),
    pages: ordered.map((page) => page.page),
  };

  return {
    protocolVersion: FRAME_DELTA_VERSION,
    docEpoch: delta.docEpoch,
    layoutEpoch: delta.layoutEpoch,
    frameEpoch: delta.frameEpoch,
    contractVersion: delta.contractVersion,
    pages: ordered,
    damagedPageIds,
    removedPageIds,
    displayList,
  };
}

const DISPLAY_PAGE_REVISION = '__betterofficePageRevision';

/**
 * Hidden in-place mutation counter for one display page. The owned delta path
 * (`applyFrameDeltaOwned`) patches primitive positions through the SAME page
 * object, so page identity alone cannot prove content equality; identity-based
 * consumers (query-facade handle adoption) must compare this revision too. The
 * property is non-enumerable, so serialization and structural reads never see
 * it.
 */
export function displayPageRevision(page: DisplayPage): number {
  return ((page as unknown as Record<string, unknown>)[DISPLAY_PAGE_REVISION] as
    | number
    | undefined) ?? 0;
}

function bumpDisplayPageRevision(page: DisplayPage): void {
  Object.defineProperty(page, DISPLAY_PAGE_REVISION, {
    value: displayPageRevision(page) + 1,
    enumerable: false,
    configurable: true,
    writable: true,
  });
}

function shiftDisplayPagePositionsOwned(
  page: DisplayPage,
  runs: readonly FramePositionShiftRun[],
  pageIndex: number
): DisplayPage {
  // primitives are mutated through this object below, whether or not a new
  // page wrapper is returned
  bumpDisplayPageRevision(page);
  let primitiveIndex = 0;
  let runIndex = 0;
  const visit = (primitives: readonly DisplayPrimitive[]): void => {
    for (const primitive of primitives) {
      while (
        runIndex < runs.length &&
        runs[runIndex]!.start + runs[runIndex]!.count <= primitiveIndex
      ) {
        runIndex += 1;
      }
      const run = runs[runIndex];
      if (run && primitiveIndex >= run.start && primitiveIndex < run.start + run.count) {
        shiftPrimitivePositionsOwned(primitive, run.changedMask, run.delta);
      }
      primitiveIndex += 1;
    }
  };
  visit(page.primitives);
  for (const area of page.noteAreas ?? []) {
    visit(area.separatorPrimitives ?? []);
    visit(area.primitives ?? []);
  }
  if (page.header) visit(page.header.primitives);
  if (page.footer) visit(page.footer.primitives);
  if (runs.some((run) => run.start + run.count > primitiveIndex)) {
    invalid('position shift range exceeds retained primitive count');
  }
  return page.pageIndex === pageIndex ? page : { ...page, pageIndex };
}

function shiftPrimitivePositionsOwned(
  primitive: DisplayPrimitive,
  changedMask: number,
  delta: number
): void {
  const shift = (
    mask: number,
    field: 'docStart' | 'docEnd' | 'fragmentDocStart' | 'fragmentDocEnd'
  ): void => {
    if ((changedMask & mask) === 0) return;
    const current = primitive[field];
    if (typeof current !== 'number') invalid(`position shift requires retained ${field}`);
    const value = current + delta;
    if (!Number.isSafeInteger(value)) invalid(`position shift overflows ${field}`);
    primitive[field] = value;
  };
  shift(POSITION_DOC_START, 'docStart');
  shift(POSITION_DOC_END, 'docEnd');
  shift(POSITION_FRAGMENT_START, 'fragmentDocStart');
  shift(POSITION_FRAGMENT_END, 'fragmentDocEnd');
  if ((changedMask & POSITION_INLINE_WIDGET) !== 0) {
    if (
      !primitive.inlineSdtWidget ||
      !Number.isSafeInteger(primitive.inlineSdtWidget.pos + delta)
    ) {
      invalid('position shift requires retained inline widget metadata');
    }
    primitive.inlineSdtWidget.pos += delta;
  }
}

function decodeStringTable(reader: BinaryReader, start: number, end: number): string[] {
  let offset = start;
  if (offset + 4 > end) invalid('truncated string table count');
  const count = reader.u32(offset);
  offset += 4;
  if (count > MAX_CONTAINER_ITEMS) invalid('string table count exceeds decoder limit');
  const decoder = new TextDecoder('utf-8', { fatal: true });
  const strings: string[] = [];
  for (let index = 0; index < count; index++) {
    if (offset + 4 > end) invalid('truncated string length');
    const length = reader.u32(offset);
    offset += 4;
    const next = checkedAdd(offset, length, 'string bytes');
    if (next > end) invalid('string exceeds string table bounds');
    try {
      strings.push(decoder.decode(reader.bytes.subarray(offset, next)));
    } catch {
      invalid('string table contains invalid UTF-8');
    }
    offset = next;
  }
  if (offset !== end) invalid('string table length/count mismatch');
  return strings;
}

function validateDataRegions(
  reader: BinaryReader,
  operations: readonly RawPageOp[],
  stringsEnd: number,
  dataOffset: number
): void {
  const regions: Array<{ start: number; end: number }> = [];
  for (const operation of operations) {
    if (operation.opcode === PAGE_OP_UPSERT) {
      if (operation.primitiveCount > 0) {
        regions.push({
          start: operation.primitiveIdsOffset,
          end: operation.primitiveIdsOffset + operation.primitiveCount * 8,
        });
      }
      regions.push({
        start: operation.payloadOffset,
        end: operation.payloadOffset + operation.payloadLength,
      });
    } else if (
      operation.opcode === PAGE_OP_PATCH_POSITIONS ||
      operation.opcode === PAGE_OP_SHIFT_POSITIONS
    ) {
      regions.push({
        start: operation.payloadOffset,
        end: operation.payloadOffset + operation.payloadLength,
      });
    }
  }
  regions.sort((left, right) => left.start - right.start || left.end - right.end);

  assertZeroPadding(reader, stringsEnd, dataOffset);
  let cursor = dataOffset;
  for (const region of regions) {
    if (region.start < cursor) invalid('data regions overlap');
    assertZeroPadding(reader, cursor, region.start);
    cursor = region.end;
  }
  if (cursor !== reader.bytes.byteLength) invalid('data section has trailing bytes');
}

function decodePositionShiftRuns(
  reader: BinaryReader,
  operation: RawPageOp
): FramePositionShiftRun[] {
  let offset = operation.payloadOffset;
  const end = operation.payloadOffset + operation.payloadLength;
  const require = (length: number, label: string): number => {
    const next = checkedAdd(offset, length, label);
    if (next > end) invalid(`truncated ${label}`);
    const current = offset;
    offset = next;
    return current;
  };
  const count = reader.u32(require(4, 'position shift run count'));
  if (count !== operation.primitiveCount || count === 0 || count > MAX_CONTAINER_ITEMS) {
    invalid('position shift run count mismatch');
  }
  if (reader.u32(require(4, 'position shift reserved word')) !== 0) {
    invalid('position shift reserved word is nonzero');
  }
  const runs: FramePositionShiftRun[] = [];
  let previousEnd = 0;
  for (let index = 0; index < count; index++) {
    const start = reader.u32(require(4, 'position shift start'));
    const runCount = reader.u32(require(4, 'position shift count'));
    const changedMask = reader.u8(require(1, 'position shift changed mask'));
    require(7, 'position shift reserved bytes');
    for (let reserved = offset - 7; reserved < offset; reserved++) {
      if (reader.u8(reserved) !== 0) invalid('position shift reserved bytes are nonzero');
    }
    const delta = reader.safeI64(require(8, 'position shift delta'), 'position shift delta');
    const runEnd = checkedAdd(start, runCount, 'position shift range');
    if (
      runCount === 0 ||
      start < previousEnd ||
      changedMask === 0 ||
      (changedMask & ~POSITION_MASK) !== 0 ||
      delta === 0
    ) {
      invalid('position shift run is invalid');
    }
    previousEnd = runEnd;
    runs.push({ start, count: runCount, changedMask, delta });
  }
  if (offset !== end) invalid('position shift byte length/count mismatch');
  return runs;
}

function assertZeroPadding(reader: BinaryReader, start: number, end: number): void {
  for (let offset = start; offset < end; offset++) {
    if (reader.u8(offset) !== 0) invalid('alignment padding is nonzero');
  }
}

function decodePositionPatches(
  reader: BinaryReader,
  operation: RawPageOp
): FramePrimitivePositionPatch[] {
  let offset = operation.payloadOffset;
  const end = operation.payloadOffset + operation.payloadLength;
  const require = (length: number, label: string): number => {
    const next = checkedAdd(offset, length, label);
    if (next > end) invalid(`truncated ${label}`);
    const current = offset;
    offset = next;
    return current;
  };

  const count = reader.u32(require(4, 'position patch count'));
  if (count !== operation.primitiveCount) invalid('position patch count mismatch');
  if (count === 0 || count > MAX_CONTAINER_ITEMS) invalid('invalid position patch count');
  if (reader.u32(require(4, 'position patch reserved word')) !== 0) {
    invalid('position patch reserved word is nonzero');
  }

  const seenIds = new Set<bigint>();
  const patches: FramePrimitivePositionPatch[] = [];
  const fields = [
    [POSITION_DOC_START, 'docStart'],
    [POSITION_DOC_END, 'docEnd'],
    [POSITION_FRAGMENT_START, 'fragmentDocStart'],
    [POSITION_FRAGMENT_END, 'fragmentDocEnd'],
    [POSITION_INLINE_WIDGET, 'inlineWidgetPos'],
  ] as const;
  for (let index = 0; index < count; index++) {
    const primitiveId = reader.u64(require(8, 'position patch primitive id'));
    if (primitiveId === 0n || seenIds.has(primitiveId)) {
      invalid('position patch carries an invalid or duplicate primitive id');
    }
    seenIds.add(primitiveId);
    const changedMask = reader.u8(require(1, 'position patch changed mask'));
    const presentMask = reader.u8(require(1, 'position patch present mask'));
    if (changedMask === 0 || (changedMask & ~POSITION_MASK) !== 0) {
      invalid('position patch changed mask is invalid');
    }
    if ((presentMask & ~changedMask) !== 0) {
      invalid('position patch present mask is not a subset of changed fields');
    }
    if (reader.u16(require(2, 'position patch reserved bytes')) !== 0) {
      invalid('position patch reserved bytes are nonzero');
    }

    const patch: {
      primitiveId: bigint;
      docStart?: number | null;
      docEnd?: number | null;
      fragmentDocStart?: number | null;
      fragmentDocEnd?: number | null;
      inlineWidgetPos?: number | null;
    } = { primitiveId };
    for (const [mask, field] of fields) {
      if ((changedMask & mask) === 0) continue;
      patch[field] =
        (presentMask & mask) === 0
          ? null
          : reader.safeI64(require(8, `position patch ${field}`), field);
    }
    patches.push(patch);
  }
  if (offset !== end) invalid('position patch byte length/count mismatch');
  return patches;
}

function patchDisplayPagePositions(
  page: DisplayPage,
  primitiveIds: BigUint64Array,
  patches: readonly FramePrimitivePositionPatch[],
  pageIndex: number
): DisplayPage {
  const byId = new Map<bigint, FramePrimitivePositionPatch>();
  for (const patch of patches) {
    if (byId.has(patch.primitiveId)) invalid(`duplicate position patch ${patch.primitiveId}`);
    byId.set(patch.primitiveId, patch);
  }

  let primitiveIndex = 0;
  const patchPrimitives = (primitives: readonly DisplayPrimitive[]): DisplayPrimitive[] =>
    primitives.map((primitive) => {
      if (primitiveIndex >= primitiveIds.length) {
        invalid('retained primitive id array is shorter than its page');
      }
      const primitiveId = primitiveIds[primitiveIndex++]!;
      const patch = byId.get(primitiveId);
      if (!patch) return primitive;
      byId.delete(primitiveId);
      return patchPrimitivePositions(primitive, patch);
    });

  const primitives = patchPrimitives(page.primitives);
  const noteAreas = page.noteAreas?.map((area) => ({
    ...area,
    ...(area.separatorPrimitives
      ? { separatorPrimitives: patchPrimitives(area.separatorPrimitives) }
      : {}),
    ...(area.primitives ? { primitives: patchPrimitives(area.primitives) } : {}),
  }));
  const header = page.header
    ? { ...page.header, primitives: patchPrimitives(page.header.primitives) }
    : undefined;
  const footer = page.footer
    ? { ...page.footer, primitives: patchPrimitives(page.footer.primitives) }
    : undefined;

  if (primitiveIndex !== primitiveIds.length) {
    invalid('retained primitive id array is longer than its page');
  }
  if (byId.size !== 0) invalid('position patch references an unknown primitive id');
  return {
    ...page,
    pageIndex,
    primitives,
    ...(noteAreas ? { noteAreas } : {}),
    ...(header ? { header } : {}),
    ...(footer ? { footer } : {}),
  };
}

function shiftDisplayPagePositions(
  page: DisplayPage,
  runs: readonly FramePositionShiftRun[],
  pageIndex: number
): DisplayPage {
  let primitiveIndex = 0;
  let runIndex = 0;
  const shiftPrimitives = (primitives: readonly DisplayPrimitive[]): DisplayPrimitive[] =>
    primitives.map((primitive) => {
      while (
        runIndex < runs.length &&
        runs[runIndex]!.start + runs[runIndex]!.count <= primitiveIndex
      ) {
        runIndex += 1;
      }
      const run = runs[runIndex];
      const shifted =
        run && primitiveIndex >= run.start && primitiveIndex < run.start + run.count
          ? shiftPrimitivePositions(primitive, run.changedMask, run.delta)
          : primitive;
      primitiveIndex += 1;
      return shifted;
    });

  const primitives = shiftPrimitives(page.primitives);
  const noteAreas = page.noteAreas?.map((area) => ({
    ...area,
    ...(area.separatorPrimitives
      ? { separatorPrimitives: shiftPrimitives(area.separatorPrimitives) }
      : {}),
    ...(area.primitives ? { primitives: shiftPrimitives(area.primitives) } : {}),
  }));
  const header = page.header
    ? { ...page.header, primitives: shiftPrimitives(page.header.primitives) }
    : undefined;
  const footer = page.footer
    ? { ...page.footer, primitives: shiftPrimitives(page.footer.primitives) }
    : undefined;
  if (runs.some((run) => run.start + run.count > primitiveIndex)) {
    invalid('position shift range exceeds retained primitive count');
  }
  return {
    ...page,
    pageIndex,
    primitives,
    ...(noteAreas ? { noteAreas } : {}),
    ...(header ? { header } : {}),
    ...(footer ? { footer } : {}),
  };
}

function shiftPrimitivePositions(
  primitive: DisplayPrimitive,
  changedMask: number,
  delta: number
): DisplayPrimitive {
  const next: DisplayPrimitive = { ...primitive };
  const shift = (
    mask: number,
    field: 'docStart' | 'docEnd' | 'fragmentDocStart' | 'fragmentDocEnd'
  ): void => {
    if ((changedMask & mask) === 0) return;
    const current = next[field];
    if (typeof current !== 'number') invalid(`position shift requires retained ${field}`);
    const value = current + delta;
    if (!Number.isSafeInteger(value)) invalid(`position shift overflows ${field}`);
    next[field] = value;
  };
  shift(POSITION_DOC_START, 'docStart');
  shift(POSITION_DOC_END, 'docEnd');
  shift(POSITION_FRAGMENT_START, 'fragmentDocStart');
  shift(POSITION_FRAGMENT_END, 'fragmentDocEnd');
  if ((changedMask & POSITION_INLINE_WIDGET) !== 0) {
    if (!next.inlineSdtWidget || !Number.isSafeInteger(next.inlineSdtWidget.pos + delta)) {
      invalid('position shift requires retained inline widget metadata');
    }
    next.inlineSdtWidget = { ...next.inlineSdtWidget, pos: next.inlineSdtWidget.pos + delta };
  }
  return next;
}

function patchPrimitivePositions(
  primitive: DisplayPrimitive,
  patch: FramePrimitivePositionPatch
): DisplayPrimitive {
  const next: DisplayPrimitive = { ...primitive };
  const setOptional = (
    field: 'docStart' | 'docEnd' | 'fragmentDocStart' | 'fragmentDocEnd',
    value: number | null
  ): void => {
    if (value === null) delete next[field];
    else next[field] = value;
  };
  if (patch.docStart !== undefined) setOptional('docStart', patch.docStart);
  if (patch.docEnd !== undefined) setOptional('docEnd', patch.docEnd);
  if (patch.fragmentDocStart !== undefined) {
    setOptional('fragmentDocStart', patch.fragmentDocStart);
  }
  if (patch.fragmentDocEnd !== undefined) setOptional('fragmentDocEnd', patch.fragmentDocEnd);
  if (patch.inlineWidgetPos !== undefined) {
    if (patch.inlineWidgetPos === null || !next.inlineSdtWidget) {
      invalid('inline widget position patch requires retained widget metadata');
    }
    next.inlineSdtWidget = { ...next.inlineSdtWidget, pos: patch.inlineWidgetPos };
  }
  return next;
}

class ValueCursor {
  offset: number;

  constructor(
    private readonly reader: BinaryReader,
    private readonly strings: readonly string[],
    offset: number,
    private readonly end: number
  ) {
    this.offset = offset;
  }

  value(depth: number): unknown {
    if (depth > MAX_VALUE_DEPTH) invalid('typed value nesting exceeds decoder limit');
    const tag = this.readU8('value tag');
    switch (tag) {
      case VALUE_NULL:
        return null;
      case VALUE_FALSE:
        return false;
      case VALUE_TRUE:
        return true;
      case VALUE_I64:
        return this.safeInteger(this.readI64('i64 value'), 'signed value');
      case VALUE_U64:
        return this.safeInteger(this.readU64('u64 value'), 'unsigned value');
      case VALUE_F64: {
        const value = this.readF64('f64 value');
        if (!Number.isFinite(value)) invalid('non-finite f64 value');
        return value;
      }
      case VALUE_STRING:
        return this.string(this.readU32('string id'));
      case VALUE_ARRAY: {
        const length = this.readU32('array byte length');
        const count = this.readU32('array item count');
        if (count > MAX_CONTAINER_ITEMS) invalid('array item count exceeds decoder limit');
        const containerEnd = checkedAdd(this.offset, length, 'array payload');
        if (containerEnd > this.end) invalid('array payload exceeds parent bounds');
        const values: unknown[] = [];
        for (let index = 0; index < count; index++) values.push(this.value(depth + 1));
        if (this.offset !== containerEnd) invalid('array byte length/count mismatch');
        return values;
      }
      case VALUE_OBJECT: {
        const length = this.readU32('object byte length');
        const count = this.readU32('object field count');
        if (count > MAX_CONTAINER_ITEMS) invalid('object field count exceeds decoder limit');
        const containerEnd = checkedAdd(this.offset, length, 'object payload');
        if (containerEnd > this.end) invalid('object payload exceeds parent bounds');
        const value: Record<string, unknown> = {};
        for (let index = 0; index < count; index++) {
          const key = this.string(this.readU32('object key id'));
          if (Object.prototype.hasOwnProperty.call(value, key))
            invalid(`duplicate object key ${key}`);
          Object.defineProperty(value, key, {
            value: this.value(depth + 1),
            enumerable: true,
            configurable: true,
            writable: true,
          });
        }
        if (this.offset !== containerEnd) invalid('object byte length/count mismatch');
        return value;
      }
      case VALUE_GLYPH_ARRAY: {
        const length = this.readU32('glyph array byte length');
        const count = this.readU32('glyph array item count');
        if (count > MAX_CONTAINER_ITEMS) invalid('glyph array item count exceeds decoder limit');
        const containerEnd = checkedAdd(this.offset, length, 'glyph array payload');
        if (containerEnd > this.end) invalid('glyph array payload exceeds parent bounds');
        const glyphs: Array<Record<string, number>> = [];
        for (let index = 0; index < count; index++) {
          const id = this.readU32('glyph id');
          const x = this.finiteF64('glyph x');
          const y = this.finiteF64('glyph y');
          const cluster = this.readU32('glyph cluster');
          const advance = this.finiteF64('glyph advance');
          const flags = this.readU8('glyph flags');
          if ((flags & ~(GLYPH_LOGICAL_ORDER | GLYPH_BIDI_LEVEL)) !== 0) {
            invalid('glyph flags contain unknown bits');
          }
          const glyph: Record<string, number> = { id, x, y, cluster, advance };
          if ((flags & GLYPH_LOGICAL_ORDER) !== 0) {
            glyph.logicalOrder = this.safeInteger(
              this.readU64('glyph logical order'),
              'glyph logical order'
            );
          }
          if ((flags & GLYPH_BIDI_LEVEL) !== 0) {
            glyph.bidiLevel = this.readU8('glyph bidi level');
          }
          glyphs.push(glyph);
        }
        if (this.offset !== containerEnd) invalid('glyph array byte length/count mismatch');
        return glyphs;
      }
      default:
        invalid(`unknown typed value opcode ${tag}`);
    }
  }

  private string(id: number): string {
    const value = this.strings[id];
    if (value === undefined) invalid(`string id ${id} is out of bounds`);
    return value;
  }

  private safeInteger(value: bigint, label: string): number {
    if (value < MIN_SAFE_BIGINT || value > MAX_SAFE_BIGINT) {
      invalid(`${label} exceeds JavaScript safe-integer range`);
    }
    return Number(value);
  }

  private require(size: number, label: string): number {
    const next = checkedAdd(this.offset, size, label);
    if (next > this.end) invalid(`truncated ${label}`);
    const current = this.offset;
    this.offset = next;
    return current;
  }

  private readU8(label: string): number {
    return this.reader.u8(this.require(1, label));
  }

  private readU32(label: string): number {
    return this.reader.u32(this.require(4, label));
  }

  private readU64(label: string): bigint {
    return this.reader.u64(this.require(8, label));
  }

  private readI64(label: string): bigint {
    return this.reader.i64(this.require(8, label));
  }

  private readF64(label: string): number {
    return this.reader.f64(this.require(8, label));
  }

  private finiteF64(label: string): number {
    const value = this.readF64(label);
    if (!Number.isFinite(value)) invalid(`non-finite ${label}`);
    return value;
  }
}

class BinaryReader {
  readonly view: DataView;

  constructor(readonly bytes: Uint8Array) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  requireLength(length: number, label: string): void {
    if (this.bytes.byteLength < length) invalid(`truncated ${label}`);
  }

  u8(offset: number): number {
    this.bounds(offset, 1);
    return this.view.getUint8(offset);
  }

  u16(offset: number): number {
    this.bounds(offset, 2);
    return this.view.getUint16(offset, true);
  }

  u32(offset: number): number {
    this.bounds(offset, 4);
    return this.view.getUint32(offset, true);
  }

  u64(offset: number): bigint {
    this.bounds(offset, 8);
    return this.view.getBigUint64(offset, true);
  }

  i64(offset: number): bigint {
    this.bounds(offset, 8);
    return this.view.getBigInt64(offset, true);
  }

  f64(offset: number): number {
    this.bounds(offset, 8);
    return this.view.getFloat64(offset, true);
  }

  safeU64(offset: number, label: string): number {
    const value = this.u64(offset);
    if (value > MAX_SAFE_BIGINT) invalid(`${label} exceeds JavaScript safe-integer range`);
    return Number(value);
  }

  safeI64(offset: number, label: string): number {
    const value = this.i64(offset);
    if (value < MIN_SAFE_BIGINT || value > MAX_SAFE_BIGINT) {
      invalid(`${label} exceeds JavaScript safe-integer range`);
    }
    return Number(value);
  }

  private bounds(offset: number, length: number): void {
    if (offset < 0 || offset + length > this.bytes.byteLength) invalid('read exceeds frame bounds');
  }
}

function validateDisplayPage(page: DisplayPage, pageIndex: number, primitiveCount: number): void {
  if (!page || typeof page !== 'object' || Array.isArray(page))
    invalid('page payload is not an object');
  if (page.pageIndex !== pageIndex) invalid('page payload/index operation mismatch');
  if (
    !Number.isFinite(page.width) ||
    page.width < 0 ||
    !Number.isFinite(page.height) ||
    page.height < 0
  ) {
    invalid('page dimensions are invalid');
  }
  if (!Array.isArray(page.primitives)) invalid('page primitives are not an array');
  let actual = page.primitives.length;
  for (const area of page.noteAreas ?? []) {
    if (!Array.isArray(area.separatorPrimitives ?? []) || !Array.isArray(area.primitives ?? [])) {
      invalid('note primitive collections are invalid');
    }
    actual += (area.separatorPrimitives?.length ?? 0) + (area.primitives?.length ?? 0);
  }
  for (const region of [page.header, page.footer]) {
    if (region && !Array.isArray(region.primitives))
      invalid('header/footer primitives are invalid');
    actual += region?.primitives.length ?? 0;
  }
  if (actual !== primitiveCount) invalid('primitive id count does not match decoded page');
}

function checkedAdd(left: number, right: number, label: string): number {
  const value = left + right;
  if (!Number.isSafeInteger(value) || value < 0) invalid(`${label} byte offset overflow`);
  return value;
}

function checkedMultiply(left: number, right: number, label: string): number {
  const value = left * right;
  if (!Number.isSafeInteger(value) || value < 0) invalid(`${label} byte length overflow`);
  return value;
}

function invalid(message: string): never {
  throw new FrameDeltaError('invalid-frame', message);
}
