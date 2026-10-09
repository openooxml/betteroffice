import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { PartsMap } from '../../docx/rezip/parts';
import {
  applyFrameDelta,
  applyFrameDeltaInternal,
  applyFrameDeltaOwned,
  decodeFrameDelta,
  decodeFrameDeltaSteps,
  displayPageNoteAnchorRevision,
  displayPageRevision,
  displayPageShiftsSince,
  FRAME_DELTA_VERSION,
  FrameDeltaError,
  retainedFramePageById,
  type DecodedFrameDelta,
  type FramePageOperation,
  type FramePageShiftRange,
  type RetainedFrame,
} from './frameDelta';
import { createDisplayListQueries } from './displayListQueries';
import type { RustDisplayListQueryEngine } from './rustDisplayList';
import type { DisplayList, DisplayPage, DisplayPrimitive } from './displayList';

const WASM = resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm');
const LAYOUT_WASM = resolve(import.meta.dir, '../../wasm/generated/layout/docx_layout_bg.wasm');
const hasEditWasm = existsSync(WASM);
const hasLayoutWasm = hasEditWasm && existsSync(LAYOUT_WASM);
let createEditSession: typeof import('../../wasm/edit').createEditSession;
let layoutWasm: typeof import('../../wasm/layout');
let rezipPartsToArrayBuffer: typeof import('../../docx/rezip/parts').rezipPartsToArrayBuffer;
let toBytes: typeof import('../../docx/rezip/parts').toBytes;
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

const PRESENT_ONLY = 1 << 7;
const ABSENT_ANCHOR = -(2n ** 63n);

it('resumable decoding matches synchronous decoding and still validates late failures', () => {
  const bytes = shiftFrame([[0, 1, 1, 1]], []);
  const steps = decodeFrameDeltaSteps(bytes);
  let yields = 0;
  let step = steps.next();
  while (!step.done) {
    yields += 1;
    step = steps.next();
  }
  expect(yields).toBeGreaterThan(0);
  expect(step.value).toEqual(decodeFrameDelta(bytes));
  const invalid = shiftFrame([[0, 1, PRESENT_ONLY, 1]], []);
  const invalidSteps = decodeFrameDeltaSteps(invalid);
  expect(() => {
    while (!invalidSteps.next().done) {}
  }).toThrow();
});

it('a recovery upsert reuses only a page with matching identity, content and primitive ids', () => {
  const previous = notedFrame();
  const page = previous.pages[0]!;
  const recovery: DecodedFrameDelta = {
    protocolVersion: FRAME_DELTA_VERSION,
    full: true,
    docEpoch: 2,
    layoutEpoch: 2,
    frameEpoch: 2,
    baseFrameEpoch: 0,
    pageCount: 1,
    bytes: new Uint8Array(),
    operations: [{ kind: 'upsert', ...page, page: structuredClone(page.page) }],
  };
  const next = applyFrameDelta(previous, recovery);
  expect(next.pages[0]).toBe(page);
  expect(next.displayList.pages[0]).toBe(page.page);
  expect([...next.damagedPageIds]).toEqual([]);
  for (const changed of [
    { fingerprint: 2n },
    { pageId: 2n },
    { primitiveIds: new BigUint64Array([2n]) },
  ]) {
    const updated = applyFrameDelta(previous, {
      ...recovery,
      operations: [{ kind: 'upsert', ...page, ...changed, page: structuredClone(page.page) }],
    });
    expect(updated.displayList.pages[0]).not.toBe(page.page);
    expect(updated.damagedPageIds.size).toBe(1);
  }
});

type CraftedRun = [start: number, count: number, mask: number, delta: number];
type CraftedAnchor = [area: number, note: number, start: bigint, end: bigint];

/** A one-page shift-positions delta laid out as the engine writes it. */
function shiftFrame(
  runs: CraftedRun[],
  anchors: CraftedAnchor[],
  options: {
    recordAnchors?: number;
    opcode?: number;
    tail?: number;
    padding?: number;
    spanDelta?: bigint;
    flags?: number;
  } = {}
): Uint8Array {
  const { recordAnchors = anchors.length, opcode = 5, tail = 0, padding = 0, spanDelta } = options;
  const dataOffset = 136;
  const payload =
    8 +
    (spanDelta !== undefined ? 8 : 0) +
    runs.length * 24 +
    (anchors.length > 0 ? 8 + anchors.length * 24 : 0) +
    padding;
  const bytes = new Uint8Array(dataOffset + payload);
  const view = new DataView(bytes.buffer);
  bytes.set([0x46, 0x44, 0x56, 0x31]);
  view.setUint16(4, FRAME_DELTA_VERSION, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, bytes.byteLength, true);
  view.setBigUint64(16, 1n, true);
  view.setBigUint64(24, 2n, true);
  view.setBigUint64(32, 2n, true);
  view.setBigUint64(40, 1n, true);
  view.setUint32(48, 1, true);
  view.setUint32(52, 1, true);
  view.setUint32(56, 80, true);
  view.setUint32(60, 128, true);
  view.setUint32(64, 4, true);
  view.setUint32(68, dataOffset, true);
  bytes[80] = opcode;
  view.setBigUint64(88, 1n, true);
  view.setBigUint64(96, 2n, true);
  view.setUint32(104, runs.length, true);
  view.setUint32(112, dataOffset, true);
  view.setUint32(116, payload, true);
  view.setUint32(120, recordAnchors, true);
  view.setUint32(124, tail, true);
  let at = dataOffset;
  view.setUint32(at, runs.length, true);
  view.setUint32(at + 4, options.flags ?? (spanDelta !== undefined ? 1 : 0), true);
  at += 8;
  if (spanDelta !== undefined) {
    view.setBigInt64(at, spanDelta, true);
    at += 8;
  }
  for (const [start, count, mask, delta] of runs) {
    view.setUint32(at, start, true);
    view.setUint32(at + 4, count, true);
    bytes[at + 8] = mask;
    view.setBigInt64(at + 16, BigInt(delta), true);
    at += 24;
  }
  if (anchors.length > 0) {
    view.setUint32(at, anchors.length, true);
    at += 8;
    for (const [area, note, start, end] of anchors) {
      view.setUint32(at, area, true);
      view.setUint32(at + 4, note, true);
      view.setBigInt64(at + 8, start, true);
      view.setBigInt64(at + 16, end, true);
      at += 24;
    }
  }
  return bytes;
}

/** A retained one-page frame whose page has a body primitive and one footnote. */
function notedFrame(): RetainedFrame {
  const page = {
    pageIndex: 0,
    width: 100,
    height: 100,
    primitives: [
      {
        kind: 'text',
        text: 'x',
        x: 10,
        baselineY: 20,
        width: 10,
        font: '400 16px Calibri',
        color: '#000000',
        docStart: 1,
        docEnd: 2,
      },
    ],
    noteAreas: [
      { kind: 'footnote', noteIds: [1], notes: [{ id: 1, anchorDocStart: 3, anchorDocEnd: 4 }] },
    ],
  } as unknown as DisplayPage;
  return {
    protocolVersion: FRAME_DELTA_VERSION,
    docEpoch: 1,
    layoutEpoch: 1,
    frameEpoch: 1,
    pages: [
      { pageIndex: 0, pageId: 1n, fingerprint: 1n, primitiveIds: new BigUint64Array([1n]), page },
    ],
    damagedPageIds: new Set(),
    removedPageIds: new Set(),
    displayList: { pages: [page] },
  };
}

/** A page record with no data region. */
type CraftedRange = {
  opcode?: number;
  pageIndex: number;
  pageId: bigint;
  count?: number;
  delta?: bigint;
  salt?: bigint;
};

/** Encode range shifts and payload-free page operations. */
function shiftRangeFrame(
  operations: readonly CraftedRange[] = [{ pageIndex: 1, pageId: 2n, count: 3, delta: -2n }],
  pageCount = 5,
  full = false
): Uint8Array {
  const stringsOffset = 80 + operations.length * 48;
  const bytes = new Uint8Array(stringsOffset + 8);
  const view = new DataView(bytes.buffer);
  bytes.set([0x46, 0x44, 0x56, 0x31]);
  view.setUint16(4, FRAME_DELTA_VERSION, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, bytes.byteLength, true);
  view.setUint32(12, full ? 1 : 0, true);
  view.setBigUint64(16, 1n, true);
  view.setBigUint64(24, 2n, true);
  view.setBigUint64(32, 2n, true);
  view.setBigUint64(40, full ? 0n : 1n, true);
  view.setUint32(48, pageCount, true);
  view.setUint32(52, operations.length, true);
  view.setUint32(56, 80, true);
  view.setUint32(60, stringsOffset, true);
  view.setUint32(64, 4, true);
  view.setUint32(68, bytes.byteLength, true);
  operations.forEach((operation, index) => {
    const at = 80 + index * 48;
    const opcode = operation.opcode ?? 6;
    bytes[at] = opcode;
    view.setUint32(at + 4, operation.pageIndex, true);
    view.setBigUint64(at + 8, operation.pageId, true);
    view.setBigUint64(at + 16, operation.salt ?? 2n, true);
    if (opcode === 6) {
      view.setUint32(at + 24, operation.count ?? 1, true);
      view.setBigInt64(at + 32, operation.delta ?? 1n, true);
    }
  });
  return bytes;
}

/** Retain hand-built pages with stable, distinct primitive identities. */
function retainedPagesFrame(
  displayPages: DisplayPage[],
  fingerprints?: readonly bigint[]
): RetainedFrame {
  const pages = displayPages.map((page, pageIndex) => {
    const count =
      page.primitives.length +
      (page.noteAreas ?? []).reduce(
        (total, area) =>
          total + (area.separatorPrimitives?.length ?? 0) + (area.primitives?.length ?? 0),
        0
      ) +
      (page.header?.primitives.length ?? 0) +
      (page.footer?.primitives.length ?? 0);
    return {
      pageIndex,
      pageId: BigInt(pageIndex + 1),
      fingerprint: fingerprints?.[pageIndex] ?? BigInt(pageIndex + 1),
      primitiveIds: BigUint64Array.from({ length: count }, (_, index) =>
        BigInt(pageIndex * 100 + index + 1)
      ),
      page,
    };
  });
  return { ...notedFrame(), pages, displayList: { pages: displayPages } };
}

/** Build a delta against a retained fixture. */
function pageDelta(
  previous: RetainedFrame,
  operations: readonly FramePageOperation[],
  overrides: Partial<DecodedFrameDelta> = {}
): DecodedFrameDelta {
  return {
    protocolVersion: FRAME_DELTA_VERSION,
    full: false,
    docEpoch: previous.docEpoch,
    layoutEpoch: previous.layoutEpoch + 1,
    frameEpoch: previous.frameEpoch + 1,
    baseFrameEpoch: previous.frameEpoch,
    pageCount: previous.pages.length,
    operations,
    bytes: new Uint8Array(),
    ...overrides,
  };
}

/** Mixed body fields, separate regions, placeholders, and empty pages. */
function rangeTestFrame(): RetainedFrame {
  const regionPrimitive: DisplayPrimitive = {
    kind: 'rect',
    x: 1,
    y: 2,
    w: 3,
    h: 4,
    fill: '#000',
    docStart: 50,
    docEnd: 60,
    fragmentDocStart: 49,
    fragmentDocEnd: 61,
    inlineSdtWidget: { kind: 'checkbox', groupId: 'region', pos: 51 },
  };
  return retainedPagesFrame(
    [
      {
        pageIndex: 0,
        width: 100,
        height: 200,
        sectionId: 'main',
        pageLabel: 'i',
        background: '#fff',
        contentBounds: { x: 10, y: 20, width: 80, height: 160 },
        positionSpan: [8, 30],
        watermarkPrimitiveCount: 1,
        primitives: [
          { kind: 'rect', x: 0, y: 1, w: 2, h: 3, fill: '#000', docStart: 10, fragmentDocEnd: 30 },
          {
            kind: 'rect',
            x: 4,
            y: 5,
            w: 6,
            h: 7,
            fill: '#000',
            docStart: 20,
            docEnd: 21,
            fragmentDocStart: 19,
            fragmentDocEnd: 22,
            inlineSdtWidget: { kind: 'checkbox', groupId: 'body', pos: 20, checked: true },
          },
          { kind: 'rect', x: 8, y: 9, w: 10, h: 11, fill: '#000', docEnd: undefined },
        ],
        noteAreas: [
          {
            kind: 'footnote',
            noteIds: [1],
            notes: [{ id: 1, anchorDocStart: 3, anchorDocEnd: 4 }],
            separatorPrimitives: [structuredClone(regionPrimitive)],
            primitives: [structuredClone(regionPrimitive)],
          },
        ],
        header: {
          kind: 'header',
          rId: 'rId1',
          y: 0,
          height: 10,
          primitives: [structuredClone(regionPrimitive)],
        },
        footer: {
          kind: 'footer',
          rId: 'rId2',
          y: 190,
          height: 10,
          primitives: [structuredClone(regionPrimitive)],
        },
      },
      {
        pageIndex: 1,
        width: 100,
        height: 200,
        primitives: [],
        unbuilt: true,
        positionSpan: [40, 80],
      },
      { pageIndex: 2, width: 100, height: 200, primitives: [], unbuilt: true },
      {
        pageIndex: 3,
        width: 100,
        height: 200,
        primitives: [],
        header: {
          kind: 'header',
          rId: 'rId3',
          y: 0,
          height: 10,
          primitives: [structuredClone(regionPrimitive)],
        },
      },
      { pageIndex: 4, width: 100, height: 200, primitives: [] },
    ],
    [1n, 2n, 0n, 0xffffffffffffffffn, 5n]
  );
}

describe('FrameDelta range shifts', () => {
  it('decodes payload-free signed range shifts and safe boundary deltas', () => {
    expect(decodeFrameDelta(shiftRangeFrame()).operations).toEqual([
      { kind: 'shift-range', pageIndex: 1, pageId: 2n, count: 3, delta: -2, salt: 2n },
    ]);
    for (const delta of [
      1n,
      -1n,
      BigInt(Number.MAX_SAFE_INTEGER),
      BigInt(Number.MIN_SAFE_INTEGER),
    ]) {
      expect(
        decodeFrameDelta(shiftRangeFrame([{ pageIndex: 0, pageId: 1n, delta }])).operations[0]
      ).toMatchObject({ delta: Number(delta) });
    }
    expect(
      decodeFrameDelta(
        shiftRangeFrame([
          { pageIndex: 3, pageId: 4n, count: 2 },
          { pageIndex: 0, pageId: 1n, count: 2 },
          { opcode: 3, pageIndex: 2, pageId: 3n },
        ])
      ).operations
    ).toHaveLength(3);
  });

  it('rejects every reserved byte and undeclared data', () => {
    for (const offset of [1, 2, 3, 28, 29, 30, 31, 40, 41, 42, 43, 44, 45, 46, 47]) {
      const bytes = shiftRangeFrame();
      bytes[80 + offset] = 1;
      expect(() => decodeFrameDelta(bytes)).toThrow(FrameDeltaError);
    }
    const original = shiftRangeFrame();
    const bytes = new Uint8Array(original.length + 8);
    bytes.set(original);
    new DataView(bytes.buffer).setUint32(8, bytes.length, true);
    expect(() => decodeFrameDelta(bytes)).toThrow('data section has trailing bytes');
  });

  it('rejects invalid range bounds, deltas, salts, identities, and full frames', () => {
    for (const operation of [
      { count: 0 },
      { pageIndex: 4, count: 2 },
      { pageIndex: 5 },
      { delta: 0n },
      { delta: 2n ** 53n },
      { delta: -(2n ** 53n) },
      { salt: 0n },
      { pageId: 0n },
    ]) {
      expect(() =>
        decodeFrameDelta(shiftRangeFrame([{ pageIndex: 1, pageId: 2n, ...operation }]))
      ).toThrow(FrameDeltaError);
    }
    expect(() =>
      decodeFrameDelta(shiftRangeFrame([{ pageIndex: 0xffffffff, pageId: 1n }], 0xffffffff))
    ).toThrow(FrameDeltaError);
    expect(() => decodeFrameDelta(shiftRangeFrame(undefined, 5, true))).toThrow(
      'full frame may contain only page upserts'
    );
    expect(() =>
      decodeFrameDelta(
        shiftRangeFrame([
          { pageIndex: 0, pageId: 1n },
          { pageIndex: 4, pageId: 1n },
        ])
      )
    ).toThrow('duplicate page operation');
  });

  it('rejects overlaps with named pages and other ranges', () => {
    for (const operations of [
      [
        { pageIndex: 1, pageId: 2n, count: 3 },
        { opcode: 3, pageIndex: 2, pageId: 3n },
      ],
      [
        { opcode: 2, pageIndex: 3, pageId: 4n },
        { pageIndex: 1, pageId: 2n, count: 3 },
      ],
      [
        { pageIndex: 2, pageId: 3n, count: 2 },
        { pageIndex: 1, pageId: 2n, count: 2 },
      ],
    ]) {
      expect(() => decodeFrameDelta(shiftRangeFrame(operations))).toThrow(FrameDeltaError);
    }
  });

  it('matches equivalent per-page shifts including revisions and replay logs', () => {
    const salt = 0x8000000000000001n;
    const fingerprints = [
      0x8000000000000000n,
      0x8000000000000003n,
      0x8000000000000001n,
      0x7ffffffffffffffen,
      5n,
    ];
    for (const apply of [applyFrameDelta, applyFrameDeltaOwned]) {
      for (const delta of [6, -4]) {
        const previous = rangeTestFrame();
        const equivalent = rangeTestFrame();
        const before = structuredClone(previous);
        const range: FramePageShiftRange = {
          kind: 'shift-range',
          pageIndex: 0,
          pageId: 1n,
          count: 4,
          delta,
          salt,
        };
        const shifts: FramePageOperation[] = equivalent.pages
          .slice(0, 4)
          .map(({ page, pageIndex, pageId }) => ({
            kind: 'shift-positions',
            pageIndex,
            pageId,
            fingerprint: fingerprints[pageIndex]!,
            runs:
              page.primitives.length === 0
                ? []
                : [
                    {
                      start: 0,
                      count: page.primitives.length,
                      changedMask: 0x1f | PRESENT_ONLY,
                      delta,
                    },
                  ],
            anchors: [],
            ...(page.positionSpan === undefined ? {} : { spanDelta: delta }),
          }));
        const next = apply(previous, pageDelta(previous, [range]));
        const expected = apply(equivalent, pageDelta(equivalent, shifts));
        expect(next).toEqual(expected);
        expect(next.pages.map((page) => page.fingerprint)).toEqual(fingerprints);
        expect(next.displayList.pages[0]!.primitives).toEqual([
          {
            ...before.pages[0]!.page.primitives[0],
            docStart: 10 + delta,
            fragmentDocEnd: 30 + delta,
          },
          {
            ...before.pages[0]!.page.primitives[1],
            docStart: 20 + delta,
            docEnd: 21 + delta,
            fragmentDocStart: 19 + delta,
            fragmentDocEnd: 22 + delta,
            inlineSdtWidget: { kind: 'checkbox', groupId: 'body', pos: 20 + delta, checked: true },
          },
          before.pages[0]!.page.primitives[2],
        ]);
        expect(next.displayList.pages[0]!.positionSpan).toEqual([8 + delta, 30 + delta]);
        expect(next.displayList.pages[1]!.positionSpan).toEqual([40 + delta, 80 + delta]);
        for (const field of ['noteAreas', 'header', 'footer'] as const) {
          expect(next.displayList.pages[0]![field]).toEqual(before.displayList.pages[0]![field]);
        }
        for (let index = 0; index < next.pages.length; index++) {
          const page = next.pages[index]!.page;
          const expectedPage = expected.pages[index]!.page;
          expect(next.pages[index]!.primitiveIds).toBe(previous.pages[index]!.primitiveIds);
          expect(displayPageRevision(page)).toBe(displayPageRevision(expectedPage));
          expect(displayPageNoteAnchorRevision(page)).toBe(
            displayPageNoteAnchorRevision(expectedPage)
          );
          expect(displayPageShiftsSince(page, 0)).toEqual(displayPageShiftsSince(expectedPage, 0));
          expect(Object.getOwnPropertyDescriptors(page)).toEqual(
            Object.getOwnPropertyDescriptors(expectedPage)
          );
          if (index < 4) {
            expect(displayPageRevision(page)).toBe(apply === applyFrameDeltaOwned ? 1 : 0);
            expect(page === previous.pages[index]!.page).toBe(apply === applyFrameDeltaOwned);
          }
        }
        if (apply === applyFrameDelta) expect(previous).toEqual(before);
        expect(next.pages[4]).toBe(previous.pages[4]);
        if (apply === applyFrameDeltaOwned) {
          const twice = apply(next, pageDelta(next, [range]));
          const expectedTwice = apply(
            expected,
            pageDelta(
              expected,
              shifts.map((shift, index) => ({
                ...shift,
                fingerprint: equivalent.pages[index]!.fingerprint,
              }))
            )
          );
          expect(twice).toEqual(expectedTwice);
          for (let index = 0; index < 4; index++) {
            expect(displayPageRevision(twice.pages[index]!.page)).toBe(2);
            expect(displayPageShiftsSince(twice.pages[index]!.page, 0)).toEqual(
              displayPageShiftsSince(expectedTwice.pages[index]!.page, 0)
            );
            expect(displayPageShiftsSince(twice.pages[index]!.page, 1)).toHaveLength(1);
          }
        }
      }
    }
  });

  it('rejects ranges in frames whose page identities or ordering change', () => {
    for (const apply of [applyFrameDelta, applyFrameDeltaOwned]) {
      const previous = rangeTestFrame();
      const range: FramePageShiftRange = {
        kind: 'shift-range',
        pageIndex: 0,
        pageId: 1n,
        count: 2,
        delta: 1,
        salt: 2n,
      };
      const before = structuredClone(previous);
      for (const delta of [
        pageDelta(previous, [range], { full: true, baseFrameEpoch: 0 }),
        pageDelta(previous, [range], { pageCount: 4 }),
        pageDelta(previous, [{ ...range, pageId: 2n }]),
        pageDelta(previous, [range, { kind: 'remove', pageIndex: 4, pageId: 5n }]),
        pageDelta(previous, [range, { kind: 'move', pageIndex: 4, pageId: 5n, fingerprint: 5n }]),
        pageDelta(previous, [range, { kind: 'upsert', ...previous.pages[4]!, pageId: 6n }]),
        pageDelta(previous, [
          range,
          { kind: 'patch-positions', pageIndex: 4, pageId: 4n, fingerprint: 6n, patches: [] },
        ]),
      ]) {
        expect(() => apply(previous, delta)).toThrow('requires an index-stable frame');
        expect(previous).toEqual(before);
      }
      for (const operations of [
        [range, { kind: 'upsert' as const, ...previous.pages[1]! }],
        [range, { ...range, pageIndex: 1, pageId: 2n }],
      ]) {
        expect(() => apply(previous, pageDelta(previous, operations))).toThrow('overlap');
        expect(previous).toEqual(before);
      }
    }
  });

  it('uses the existing span and primitive overflow checks', () => {
    for (const apply of [applyFrameDelta, applyFrameDeltaOwned]) {
      const range: FramePageShiftRange = {
        kind: 'shift-range',
        pageIndex: 0,
        pageId: 1n,
        count: 1,
        delta: 1,
        salt: 2n,
      };
      const span = rangeTestFrame();
      span.pages[0]!.page.positionSpan = [0, Number.MAX_SAFE_INTEGER];
      expect(() => apply(span, pageDelta(span, [range]))).toThrow('overflows positionSpan');
      for (const field of ['docStart', 'docEnd', 'fragmentDocStart', 'fragmentDocEnd'] as const) {
        const previous = rangeTestFrame();
        previous.pages[0]!.page.primitives[0]![field] = Number.MAX_SAFE_INTEGER;
        expect(() => apply(previous, pageDelta(previous, [range]))).toThrow(`overflows ${field}`);
      }
      const widget = rangeTestFrame();
      widget.pages[0]!.page.primitives[1]!.inlineSdtWidget!.pos = Number.MAX_SAFE_INTEGER;
      expect(() => apply(widget, pageDelta(widget, [range]))).toThrow(
        'requires retained inline widget metadata'
      );
    }
  });
});

describe('FrameDelta index-stable apply', () => {
  it('matches general storage for upserts, patches, shifts and ranges without changing old arrays', () => {
    for (const apply of [applyFrameDelta, applyFrameDeltaOwned]) {
      for (const contractVersion of [undefined, 7]) {
        const previous = retainedPagesFrame(
          Array.from({ length: 7 }, (_, pageIndex) => ({
            ...structuredClone(notedFrame().pages[0]!.page),
            pageIndex,
          }))
        );
        const general = structuredClone(previous);
        const oldPages = previous.pages;
        const oldListPages = previous.displayList.pages;
        const oldPageEntries = oldPages.slice();
        const oldListEntries = oldListPages.slice();
        const operations: FramePageOperation[] = [
          { kind: 'upsert', ...previous.pages[0]!, page: structuredClone(previous.pages[0]!.page) },
          {
            kind: 'upsert',
            ...previous.pages[1]!,
            fingerprint: 20n,
            primitiveIds: new BigUint64Array([200n]),
            page: { ...previous.pages[1]!.page, width: 150 },
          },
          {
            kind: 'patch-positions',
            pageIndex: 2,
            pageId: 3n,
            fingerprint: 30n,
            patches: [{ primitiveId: 201n, docStart: 12, docEnd: null }],
          },
          {
            kind: 'shift-positions',
            pageIndex: 3,
            pageId: 4n,
            fingerprint: 40n,
            runs: [{ start: 0, count: 1, changedMask: 3, delta: 2 }],
            anchors: [{ area: 0, note: 0, start: 8, end: null }],
          },
          {
            kind: 'shift-range',
            pageIndex: 4,
            pageId: 5n,
            count: 2,
            delta: -1,
            salt: 0x8000000000000001n,
          },
        ];
        retainedFramePageById(previous, 1n);
        const delta = pageDelta(previous, operations, { contractVersion });
        const next = apply(previous, delta);
        const expected = applyFrameDeltaInternal(
          general,
          delta,
          apply === applyFrameDeltaOwned,
          false
        );
        expect(next).toEqual(expected);
        expect(next).not.toBe(previous);
        expect(next.displayList).not.toBe(previous.displayList);
        expect(next.pages).not.toBe(oldPages);
        expect(next.displayList.pages).not.toBe(oldListPages);
        expect(previous.pages).toBe(oldPages);
        expect(previous.displayList.pages).toBe(oldListPages);
        expect(next.pages[0]).toBe(oldPages[0]);
        expect([...next.damagedPageIds]).toEqual([2n]);
        expect([...next.removedPageIds]).toEqual([]);
        for (let index = 0; index < next.pages.length; index++) {
          expect(oldPages[index]).toBe(oldPageEntries[index]);
          expect(oldListPages[index]).toBe(oldListEntries[index]);
          expect(displayPageRevision(next.pages[index]!.page)).toBe(
            displayPageRevision(expected.pages[index]!.page)
          );
          expect(displayPageNoteAnchorRevision(next.pages[index]!.page)).toBe(
            displayPageNoteAnchorRevision(expected.pages[index]!.page)
          );
          expect(displayPageShiftsSince(next.pages[index]!.page, 0)).toEqual(
            displayPageShiftsSince(expected.pages[index]!.page, 0)
          );
        }
      }
    }
  });

  it('shares its lazy index across stable frames and rebuilds it after moves and removals', () => {
    const get = spyOn(Map.prototype, 'get');
    try {
      const previous = rangeTestFrame();
      expect(retainedFramePageById(previous, 1n)).toBe(previous.pages[0]);
      const originalIndex = get.mock.contexts.at(-1);
      const next = applyFrameDelta(previous, pageDelta(previous, []));
      expect(retainedFramePageById(next, 1n)).toBe(next.pages[0]);
      expect(get.mock.contexts.at(-1)).toBe(originalIndex);
      const moved = applyFrameDelta(
        next,
        pageDelta(next, [
          { kind: 'move', pageIndex: 1, pageId: 1n, fingerprint: 1n },
          { kind: 'move', pageIndex: 0, pageId: 2n, fingerprint: 2n },
        ])
      );
      expect(retainedFramePageById(moved, 1n)).toBe(moved.pages[1]);
      const movedIndex = get.mock.contexts.at(-1);
      expect(movedIndex).not.toBe(originalIndex);
      const removed = applyFrameDelta(
        moved,
        pageDelta(moved, [{ kind: 'remove', pageIndex: 4, pageId: 5n }], { pageCount: 4 })
      );
      expect(retainedFramePageById(removed, 5n)).toBeUndefined();
      expect(get.mock.contexts.at(-1)).not.toBe(movedIndex);
      expect(retainedFramePageById(removed, 1n)).toBe(removed.pages[1]);
      expect(retainedFramePageById(previous, 1n)).toBe(previous.pages[0]);
      expect(retainedFramePageById(moved, 5n)).toBe(moved.pages[4]);
      expect([...removed.removedPageIds]).toEqual([5n]);
    } finally {
      get.mockRestore();
    }
  });

  it('avoids page map writes and sorting on a warm range shift', () => {
    const previous = rangeTestFrame();
    retainedFramePageById(previous, 1n);
    const set = spyOn(Map.prototype, 'set');
    const sort = spyOn(Array.prototype, 'sort');
    try {
      const next = applyFrameDeltaOwned(
        previous,
        pageDelta(previous, [
          { kind: 'shift-range', pageIndex: 0, pageId: 1n, count: 4, delta: 1, salt: 2n },
        ])
      );
      expect(next.pages).toHaveLength(5);
      expect(set).not.toHaveBeenCalled();
      expect(sort).not.toHaveBeenCalled();
    } finally {
      set.mockRestore();
      sort.mockRestore();
    }
  });

  it('keeps general-path page count and contiguity validation', () => {
    const previous = rangeTestFrame();
    expect(() =>
      applyFrameDelta(previous, pageDelta(previous, [{ kind: 'remove', pageIndex: 4, pageId: 5n }]))
    ).toThrow('applied page count');
    expect(() =>
      applyFrameDelta(
        previous,
        pageDelta(previous, [{ kind: 'move', pageIndex: 1, pageId: 1n, fingerprint: 1n }])
      )
    ).toThrow('not a contiguous zero-based sequence');
  });
});

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const FOOTNOTES = [1, 2, 3, 4];
const ENDNOTES = [1, 2];

/**
 * Twelve paragraphs on small pages: footnotes on pages spread through the
 * document, and both endnotes on the last page beside its own footnote.
 */
function notedDocx(): Uint8Array {
  const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const reference = (kind: 'footnote' | 'endnote', id: number) =>
    `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`;
  const references: Record<number, string> = {
    1: reference('footnote', 1),
    2: reference('endnote', 1),
    5: reference('footnote', 2),
    9: reference('footnote', 3),
    10: reference('endnote', 2),
    11: reference('footnote', 4),
  };
  const body = Array.from({ length: 12 }, (_, index) => {
    const words = Array.from({ length: 40 }, (_, word) => `w${index}_${word}`).join(' ');
    return `<w:p>${run(words)}${references[index] ?? ''}</w:p>`;
  }).join('');
  const notes = (kind: 'footnote' | 'endnote', ids: number[]) =>
    `<w:${kind}s ${W_NS}>` +
    `<w:${kind} w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
    ids
      .map((id) => `<w:${kind} w:id="${id}"><w:p>${run(`${kind} ${id}`)}</w:p></w:${kind}>`)
      .join('') +
    `</w:${kind}s>`;
  const relationship = (id: string, type: string, target: string) =>
    `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  const relationships = (entries: string) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
  const override = (part: string, type: string) =>
    `<Override PartName="/word/${part}.xml" ContentType="${OFFICE_DOC}.wordprocessingml.${type}+xml"/>`;
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        override('document', 'document.main') +
        override('footnotes', 'footnotes') +
        override('endnotes', 'endnotes') +
        `</Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(relationships(relationship('rId1', 'officeDocument', 'word/document.xml')))
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      relationships(
        relationship('rId2', 'footnotes', 'footnotes.xml') +
          relationship('rId3', 'endnotes', 'endnotes.xml')
      )
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document ${W_NS}><w:body>${body}<w:sectPr><w:pgSz w:w="4320" w:h="4320"/>` +
        `<w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr></w:body></w:document>`
    )
  );
  parts.set('word/footnotes.xml', toBytes(notes('footnote', FOOTNOTES)));
  parts.set('word/endnotes.xml', toBytes(notes('endnote', ENDNOTES)));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** The first shift run's mask in `frame`, rewritten, must fail to decode when invalid. */
function expectStrictShiftMasks(frame: Uint8Array): void {
  const view = new DataView(frame.buffer, frame.byteOffset, frame.byteLength);
  const operations = view.getUint32(52, true);
  let maskOffset = -1;
  for (let index = 0; index < operations && maskOffset < 0; index++) {
    const record = 80 + index * 48;
    if (frame[record] === 5) maskOffset = view.getUint32(record + 32, true) + 16;
  }
  if (maskOffset < 0) {
    expect(decodeFrameDelta(frame).operations.some((operation) => operation.kind === 'shift-range')).toBe(true);
    return;
  }
  expect(maskOffset).toBeGreaterThan(0);
  for (const mask of [PRESENT_ONLY, (frame[maskOffset]! & 0x1f) | 0x40]) {
    const patched = frame.slice();
    patched[maskOffset] = mask;
    expect(() => decodeFrameDelta(patched)).toThrow('position shift run is invalid');
  }
}

describe('FrameDelta wire round-trip', () => {
  beforeAll(async () => {
    if (!hasEditWasm) return;
    const editWasm = await import('../../wasm/edit');
    createEditSession = editWasm.createEditSession;
    await editWasm.preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  });

  it.skipIf(!hasEditWasm)('decodes wasm-encoded full and delta frames to the equivalent JSON list', () => {
    const session = createEditSession(11);
    const { paraId } = JSON.parse(session.create_story('body', 'Hello frame', 'Normal', 'left'));
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const request = JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: {
        fontChains: { 'calibri|0|0': [fontId] },
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: true,
      },
      renderEnv: {},
    });

    const envelopeFor = (): string => {
      const output = JSON.parse(session.layout_document_with_regions_json(request)) as {
        measured: unknown;
        options: unknown;
        layout: unknown;
      };
      return JSON.stringify({
        measured: output.measured,
        options: output.options,
        layout: output.layout,
        fontChains: { 'calibri|0|0': [fontId] },
        headersFooters: {
          variants: [],
          watermark: { kind: 'text', text: 'DRAFT', font: 'Calibri' },
        },
      });
    };

    const first = envelopeFor();
    const jsonList = JSON.parse(session.build_display_list_json(first)) as DisplayList;
    const fullFrame = session.build_display_list_frame(first, 0);
    const retained = applyFrameDelta(null, decodeFrameDelta(fullFrame));
    expect(retained.displayList).toEqual(jsonList);
    expect(retained.displayList.pages[0].watermarkPrimitiveCount).toBe(1);

    session.insert_text('body', paraId, 5, ' typed', undefined, undefined);
    const second = envelopeFor();
    const nextJsonList = JSON.parse(session.build_display_list_json(second)) as DisplayList;
    const deltaFrame = session.build_display_list_frame(second, retained.frameEpoch);
    const next = applyFrameDelta(retained, decodeFrameDelta(deltaFrame));
    expect(next.displayList).toEqual(nextJsonList);
    const pageText = next.displayList.pages[0].primitives
      .map((primitive) => ('text' in primitive ? (primitive.text ?? '') : ''))
      .join('');
    expect(pageText).toContain('typed');
  });

  it.skipIf(!hasEditWasm)('retains owned primitive ids, so no retained page keeps its frame buffer alive', () => {
    const session = createEditSession(12);
    session.create_story('body', 'Hello frame', 'Normal', 'left');
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const output = JSON.parse(
      session.layout_document_with_regions_json(
        JSON.stringify({
          bodyStory: 'body',
          regions: { sections: [{ sectionId: 'main', properties: {} }] },
          measurement: {
            fontChains: { 'calibri|0|0': [fontId] },
            defaults: { fontSize: 11, fontFamily: 'Calibri' },
            authoritativeShaping: true,
          },
          renderEnv: {},
        })
      )
    ) as { measured: unknown; options: unknown; layout: unknown };
    const frame = session.build_display_list_frame(
      JSON.stringify({ ...output, fontChains: { 'calibri|0|0': [fontId] } }),
      0
    );
    const delta = decodeFrameDelta(frame);
    const upserted = delta.operations.flatMap((operation) =>
      operation.kind === 'upsert' ? [[...operation.primitiveIds]] : []
    );
    const retained = [applyFrameDelta(null, delta), applyFrameDeltaOwned(null, delta)];
    // Detaching the frame buffer empties every view into it.
    structuredClone(frame.buffer, { transfer: [frame.buffer] });
    expect(frame.byteLength).toBe(0);
    expect(upserted.length).toBeGreaterThan(0);
    for (const frameAfter of retained) {
      expect(frameAfter.pages.map((page) => [...page.primitiveIds])).toEqual(upserted);
    }
  });

  it.skipIf(!hasEditWasm)('records owned position shifts and ships them as query-store shift ops', () => {
    const session = createEditSession(13);
    const sentence = 'shift the following pages with enough text to fill several tiny pages. ';
    const { paraId } = JSON.parse(
      session.create_story('body', `start me. ${sentence.repeat(12)}`, 'Normal', 'left')
    );
    // split into paragraphs so trailing pages hold untouched blocks whose doc
    // positions merely shift when the first paragraph grows
    let splitId = paraId as string;
    for (let i = 0; i < 10; i++) {
      const receipt = JSON.parse(session.split_paragraph('body', splitId, 60, undefined, undefined)) as {
        secondParaId: string;
      };
      splitId = receipt.secondParaId;
    }
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const request = JSON.stringify({
      bodyStory: 'body',
      regions: {
        sections: [
          {
            sectionId: 'main',
            properties: {
              pageWidth: 4320,
              pageHeight: 2880,
              marginTop: 300,
              marginRight: 300,
              marginBottom: 300,
              marginLeft: 300,
            },
          },
        ],
      },
      measurement: {
        fontChains: { 'calibri|0|0': [fontId] },
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: true,
      },
      renderEnv: {},
    });
    const envelopeFor = (): string => {
      const output = JSON.parse(session.layout_document_with_regions_json(request)) as {
        measured: unknown;
        options: unknown;
        layout: unknown;
      };
      return JSON.stringify({
        measured: output.measured,
        options: output.options,
        layout: output.layout,
        fontChains: { 'calibri|0|0': [fontId] },
      });
    };

    const first = envelopeFor();
    const retained = applyFrameDeltaOwned(null, decodeFrameDelta(session.build_display_list_frame(first, 0)));
    expect(retained.displayList.pages.length).toBeGreaterThan(1);
    const trailingPage = retained.displayList.pages.at(-1)!;
    const revisionBefore = displayPageRevision(trailingPage);

    const updates: string[] = [];
    let nextHandle = 1;
    const engine: RustDisplayListQueryEngine = {
      hitTestRegionsJson: () => 'null',
      verticalMoveJson: () => 'null',
      rangeRectsJson: () => '[]',
      hasDisplayListSession: () => true,
      openDisplayList: () => nextHandle++,
      closeDisplayList: () => {},
      updateDisplayList: (_handle, update) => {
        updates.push(update);
      },
      hasDisplayListUpdate: () => true,
      rangeRectsByHandle: () => '[]',
      verticalMoveByHandle: () => 'null',
    };
    const firstQueries = createDisplayListQueries(retained.displayList, engine);
    // parse every page into the store, so adoption has retained pages to shift
    firstQueries.rangeRects(0, Number.MAX_SAFE_INTEGER);
    expect(updates.length).toBe(1);

    session.insert_text('body', paraId, 5, 'x', undefined, undefined);
    const second = envelopeFor();
    const deltaFrame = session.build_display_list_frame(second, retained.frameEpoch);
    expectStrictShiftMasks(deltaFrame);
    const next = applyFrameDeltaOwned(retained, decodeFrameDelta(deltaFrame));

    // trailing pages absorb the insert as an in-place position shift with a
    // recorded, replayable run log
    expect(next.displayList.pages.at(-1)).toBe(trailingPage);
    expect(displayPageRevision(trailingPage)).toBe(revisionBefore + 1);
    const shifts = displayPageShiftsSince(trailingPage, revisionBefore);
    expect(shifts).not.toBeNull();
    expect(shifts!.length).toBe(1);
    // the whole page moves as one run over the fields each primitive has
    expect(shifts![0]!.runs).toHaveLength(1);
    expect(shifts![0]!.runs[0]!.changedMask & PRESENT_ONLY).toBe(PRESENT_ONLY);
    expect(displayPageShiftsSince(trailingPage, revisionBefore + 1)).toEqual([]);

    // handle adoption ships those shifts as compact ops instead of replacing
    // the page's serialized payload
    const secondQueries = createDisplayListQueries(next.displayList, engine, firstQueries);
    secondQueries.prime();
    expect(updates.length).toBe(2);
    const update = JSON.parse(updates[1]!) as {
      total: number;
      replace?: Array<[number, unknown]>;
      shift?: Array<[number, number, number[][][]]>;
    };
    const trailingIndex = next.displayList.pages.length - 1;
    expect(update.shift?.some(([to]) => to === trailingIndex)).toBe(true);
    expect(update.replace?.some(([to]) => to === trailingIndex)).toBeFalsy();
  });

  it('records owned shifts only after every run applies', () => {
    const page: DisplayPage = {
      pageIndex: 0,
      width: 100,
      height: 100,
      primitives: [
        {
          kind: 'text',
          text: 'x',
          x: 10,
          baselineY: 20,
          width: 10,
          font: '400 16px Calibri',
          color: '#000000',
        },
      ],
    };
    const previous: RetainedFrame = {
      protocolVersion: FRAME_DELTA_VERSION,
      docEpoch: 1,
      layoutEpoch: 1,
      frameEpoch: 1,
      pages: [
        {
          pageIndex: 0,
          pageId: 1n,
          fingerprint: 1n,
          primitiveIds: new BigUint64Array([1n]),
          page,
        },
      ],
      damagedPageIds: new Set(),
      removedPageIds: new Set(),
      displayList: { pages: [page] },
    };
    const delta: DecodedFrameDelta = {
      protocolVersion: FRAME_DELTA_VERSION,
      full: false,
      docEpoch: 1,
      layoutEpoch: 2,
      frameEpoch: 2,
      baseFrameEpoch: 1,
      pageCount: 1,
      operations: [
        {
          kind: 'shift-positions',
          pageIndex: 0,
          pageId: 1n,
          fingerprint: 2n,
          runs: [{ start: 0, count: 1, changedMask: 1, delta: 1 }],
          anchors: [],
        },
      ],
      bytes: new Uint8Array(),
    };

    expect(() => applyFrameDeltaOwned(previous, delta)).toThrow('requires retained docStart');
    expect(displayPageRevision(page)).toBe(0);
    expect(displayPageShiftsSince(page, 0)).toEqual([]);
  });

  it('moves a page the same with present-only and exact shift runs', () => {
    const text = (start: number, extra: Record<string, unknown> = {}) => ({
      kind: 'text' as const,
      text: 'x',
      x: 10,
      baselineY: 20,
      width: 10,
      font: '400 16px Calibri',
      color: '#000000',
      docStart: start,
      docEnd: start + 1,
      ...extra,
    });
    const page = (): DisplayPage =>
      ({
        pageIndex: 0,
        width: 100,
        height: 100,
        primitives: [
          text(10),
          {
            kind: 'rect',
            x: 0,
            y: 0,
            w: 5,
            h: 5,
            fill: '#000',
            fragmentDocStart: 9,
            fragmentDocEnd: 30,
          },
          { kind: 'rect', x: 0, y: 9, w: 5, h: 5, fill: '#000' },
          text(20, { inlineSdtWidget: { kind: 'checkbox', groupId: 'g', pos: 20 } }),
        ],
        noteAreas: [
          {
            kind: 'footnote',
            separatorPrimitives: [{ kind: 'rect', x: 0, y: 90, w: 50, h: 1, fill: '#000' }],
            primitives: [text(40)],
          },
        ],
        header: { kind: 'header', rId: 'rId1', y: 0, height: 10, primitives: [text(3)] },
        footer: { kind: 'footer', rId: 'rId2', y: 90, height: 10, primitives: [text(4)] },
      }) as unknown as DisplayPage;
    const frameWith = (displayPage: DisplayPage): RetainedFrame => ({
      protocolVersion: FRAME_DELTA_VERSION,
      docEpoch: 1,
      layoutEpoch: 1,
      frameEpoch: 1,
      pages: [
        {
          pageIndex: 0,
          pageId: 1n,
          fingerprint: 1n,
          primitiveIds: new BigUint64Array([1n, 2n, 3n, 4n, 5n, 6n, 7n, 8n]),
          page: displayPage,
        },
      ],
      damagedPageIds: new Set(),
      removedPageIds: new Set(),
      displayList: { pages: [displayPage] },
    });
    type Run = { start: number; count: number; changedMask: number; delta: number };
    const shiftedBy = (runs: Run[]) =>
      ({
        protocolVersion: FRAME_DELTA_VERSION,
        full: false,
        docEpoch: 1,
        layoutEpoch: 2,
        frameEpoch: 2,
        baseFrameEpoch: 1,
        pageCount: 1,
        operations: [
          { kind: 'shift-positions', pageIndex: 0, pageId: 1n, fingerprint: 2n, runs, anchors: [] },
        ],
        bytes: new Uint8Array(),
      }) as DecodedFrameDelta;
    for (const delta of [6, -4]) {
      const exact = shiftedBy([
        { start: 0, count: 1, changedMask: 0b11, delta },
        { start: 1, count: 1, changedMask: 0b1100, delta },
        { start: 3, count: 1, changedMask: 0b10011, delta },
        { start: 5, count: 1, changedMask: 0b11, delta },
      ]);
      const presentOnly = shiftedBy([
        { start: 0, count: 4, changedMask: 0b11111 | PRESENT_ONLY, delta },
        { start: 4, count: 2, changedMask: 0b11 | PRESENT_ONLY, delta },
      ]);
      const viaExact = applyFrameDelta(frameWith(page()), exact).displayList;
      expect(applyFrameDelta(frameWith(page()), presentOnly).displayList).toEqual(viaExact);
      expect(applyFrameDeltaOwned(frameWith(page()), presentOnly).displayList).toEqual(
        applyFrameDeltaOwned(frameWith(page()), exact).displayList
      );
      expect(viaExact.pages[0]!.header!.primitives[0]).toMatchObject({ docStart: 3 });
    }
  });
});

describe('FrameDelta position span shifts', () => {
  const placeholderFrame = (span?: [number, number]): RetainedFrame => {
    const page: DisplayPage = {
      pageIndex: 0,
      width: 100,
      height: 100,
      primitives: [],
      unbuilt: true,
      ...(span ? { positionSpan: span } : {}),
    };
    return {
      ...notedFrame(),
      pages: [
        { pageIndex: 0, pageId: 1n, fingerprint: 1n, primitiveIds: new BigUint64Array(), page },
      ],
      displayList: { pages: [page] },
    };
  };

  it('applies span-only shifts like equivalent upserts with both appliers', () => {
    for (const spanDelta of [6n, -4n]) {
      const delta = decodeFrameDelta(shiftFrame([], [], { spanDelta }));
      expect(delta.operations[0]).toMatchObject({
        runs: [],
        anchors: [],
        spanDelta: Number(spanDelta),
      });
      const page = placeholderFrame([8 + Number(spanDelta), 30 + Number(spanDelta)]).pages[0]!;
      const upsert: DecodedFrameDelta = {
        ...delta,
        operations: [{ kind: 'upsert', ...page, fingerprint: 2n }],
      };
      const expected = applyFrameDelta(placeholderFrame([8, 30]), upsert).displayList;
      const previous = placeholderFrame([8, 30]);
      const copied = applyFrameDelta(previous, delta);
      expect(copied.displayList).toEqual(expected);
      expect(previous.displayList.pages[0]!.positionSpan).toEqual([8, 30]);
      const owned = applyFrameDeltaOwned(placeholderFrame([8, 30]), delta);
      expect(owned.displayList).toEqual(expected);
      expect(owned.pages[0]!.fingerprint).toBe(2n);
      expect(displayPageShiftsSince(owned.displayList.pages[0]!, 0)).toEqual([
        { runs: [], anchors: [], spanDelta: Number(spanDelta) },
      ]);
    }
  });

  it('refreshes query-store pages after owned span shifts', () => {
    const initial = placeholderFrame([8, 30]);
    const unchanged: DisplayPage = { pageIndex: 1, width: 100, height: 100, primitives: [] };
    const previous: RetainedFrame = {
      ...initial,
      pages: [
        ...initial.pages,
        {
          pageIndex: 1,
          pageId: 2n,
          fingerprint: 1n,
          primitiveIds: new BigUint64Array(),
          page: unchanged,
        },
      ],
      displayList: { pages: [...initial.displayList.pages, unchanged] },
    };
    const updates: string[] = [];
    let canUpdate = false;
    const engine: RustDisplayListQueryEngine = {
      hitTestRegionsJson: () => 'null',
      rangeRectsJson: () => '[]',
      hasDisplayListSession: () => true,
      openDisplayList: () => 1,
      closeDisplayList: () => {},
      hasDisplayListUpdate: () => canUpdate,
      updateDisplayList: (_handle, update) => {
        updates.push(update);
      },
      rangeRectsByHandle: () => '[]',
      verticalMoveByHandle: () => 'null',
    };
    const first = createDisplayListQueries(previous.displayList, engine);
    first.prime();
    canUpdate = true;
    const shifted = applyFrameDeltaOwned(
      placeholderFrame([8, 30]),
      decodeFrameDelta(shiftFrame([], [], { spanDelta: 6n }))
    );
    const original = previous.displayList.pages[0]!;
    const applied = applyFrameDeltaOwned(previous, {
      ...decodeFrameDelta(shiftFrame([], [], { spanDelta: 6n })),
      pageCount: 2,
    });
    const second = createDisplayListQueries(applied.displayList, engine, first);
    second.prime();
    expect(applied.displayList.pages[0]).toBe(original);
    expect(JSON.parse(updates[0]!)).toEqual({
      total: 2,
      reuse: [[1, 1]],
      replace: [[0, { pageIndex: 0, width: 100, height: 100, primitives: [] }]],
    });
    const fresh = createDisplayListQueries(
      { pages: [shifted.displayList.pages[0]!, unchanged] },
      engine
    );
    for (const pos of [8, 14, 30, 36, 37]) {
      expect(second.caretRect(pos)).toEqual(fresh.caretRect(pos));
    }
    first.dispose();
    second.dispose();
    fresh.dispose();
  });

  it('rejects malformed span shifts with both appliers', () => {
    const delta = decodeFrameDelta(shiftFrame([], [], { spanDelta: 1n }));
    for (const apply of [applyFrameDelta, applyFrameDeltaOwned]) {
      expect(() => apply(placeholderFrame(), delta)).toThrow('requires retained positionSpan');
      const previous = placeholderFrame([8, Number.MAX_SAFE_INTEGER]);
      expect(() => apply(previous, delta)).toThrow('overflows positionSpan');
      expect(previous.displayList.pages[0]!.positionSpan).toEqual([8, Number.MAX_SAFE_INTEGER]);
      expect(displayPageRevision(previous.displayList.pages[0]!)).toBe(0);
      expect(() => apply(placeholderFrame([1.5, 8]), delta)).toThrow('overflows positionSpan');
    }
    expect(() => decodeFrameDelta(shiftFrame([], [], { spanDelta: 2n ** 53n }))).toThrow();
    expect(() => decodeFrameDelta(shiftFrame([], [], { spanDelta: 0n }))).toThrow();
    expect(() => decodeFrameDelta(shiftFrame([], [], { spanDelta: 1n, flags: 3 }))).toThrow();
    expect(() => decodeFrameDelta(shiftFrame([], [], { flags: 1 }))).toThrow();
    expect(() => decodeFrameDelta(shiftFrame([], []))).toThrow('position shift run count mismatch');
  });
});

describe('FrameDelta note anchor shifts', () => {
  beforeAll(async () => {
    if (!hasLayoutWasm) return;
    const editWasm = await import('../../wasm/edit');
    createEditSession = editWasm.createEditSession;
    layoutWasm = await import('../../wasm/layout');
    ({ rezipPartsToArrayBuffer, toBytes } = await import('../../docx/rezip/parts'));
    await editWasm.preloadEditWasm(new Uint8Array(readFileSync(WASM)));
    await layoutWasm.preloadLayoutWasm();
  });

  it('decodes anchor sections strictly', () => {
    const run: CraftedRun = [0, 1, 0b11, 1];
    expect(() => decodeFrameDelta(shiftFrame([], []))).toThrow('position shift run count mismatch');
    expect(decodeFrameDelta(shiftFrame([], [[0, 1, 5n, ABSENT_ANCHOR]])).operations).toEqual([
      {
        kind: 'shift-positions',
        pageIndex: 0,
        pageId: 1n,
        fingerprint: 2n,
        runs: [],
        anchors: [{ area: 0, note: 1, start: 5, end: null }],
      },
    ]);
    expect(() =>
      decodeFrameDelta(shiftFrame([run], [[0, 0, 5n, 6n]], { recordAnchors: 2 }))
    ).toThrow('note anchor count mismatch');
    expect(() =>
      decodeFrameDelta(shiftFrame([run], [[0, 0, 5n, 6n]], { recordAnchors: 0 }))
    ).toThrow('position shift byte length/count mismatch');
    expect(() => decodeFrameDelta(shiftFrame([run], [[0, 0, 5n, 6n]], { padding: 8 }))).toThrow(
      'position shift byte length/count mismatch'
    );
    expect(() => decodeFrameDelta(shiftFrame([run], [[0, 0, 5n, 6n]], { tail: 1 }))).toThrow(
      'page operation tail is nonzero'
    );
    expect(() => decodeFrameDelta(shiftFrame([run], [], { opcode: 4, recordAnchors: 1 }))).toThrow(
      'page operation tail is nonzero'
    );
    expect(() =>
      decodeFrameDelta(
        shiftFrame(
          [],
          [
            [0, 1, 5n, 6n],
            [0, 1, 7n, 8n],
          ]
        )
      )
    ).toThrow('note anchors are not in strictly increasing order');
    expect(() => decodeFrameDelta(shiftFrame([], [[0, 0, 2n ** 60n, 6n]]))).toThrow(
      'safe-integer range'
    );
    const farArea = 2 ** 21;
    expect(
      decodeFrameDelta(
        shiftFrame(
          [],
          [
            [farArea, 0, 5n, 6n],
            [farArea, 1, 7n, 8n],
          ]
        )
      ).operations
    ).toHaveLength(1);
  });

  it('sets anchors after runs and rejects a missing note before touching the page', () => {
    for (const apply of [applyFrameDelta, applyFrameDeltaOwned]) {
      const moved = apply(
        notedFrame(),
        decodeFrameDelta(shiftFrame([[0, 1, 0b11, 2]], [[0, 0, 7n, ABSENT_ANCHOR]]))
      ).displayList.pages[0]!;
      expect(moved.primitives[0]).toMatchObject({ docStart: 3, docEnd: 4 });
      expect(moved.noteAreas![0]!.notes).toEqual([{ id: 1, anchorDocStart: 7 }]);
      if (apply === applyFrameDeltaOwned) expect(displayPageNoteAnchorRevision(moved)).toBe(1);

      for (const missing of [
        [1, 0, 7n, 8n],
        [0, 1, 7n, 8n],
      ] as CraftedAnchor[]) {
        const frame = notedFrame();
        const page = frame.pages[0]!.page;
        expect(() =>
          apply(frame, decodeFrameDelta(shiftFrame([[0, 1, 0b11, 2]], [missing])))
        ).toThrow('note anchor shift references an unknown note');
        expect(page.primitives[0]).toMatchObject({ docStart: 1, docEnd: 2 });
        expect(page.noteAreas![0]!.notes).toEqual([{ id: 1, anchorDocStart: 3, anchorDocEnd: 4 }]);
        expect(displayPageRevision(page)).toBe(0);
      }
    }
  });

  it.skipIf(!hasLayoutWasm)('matches a fresh layout on footnote, endnote and mixed pages, in frames and the query store', () => {
    const session = createEditSession(71);
    session.seed_from_docx(notedDocx(), undefined);
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const base = {
      bodyStory: 'body',
      regions: {
        sections: [
          {
            properties: {
              pageWidth: 4320,
              pageHeight: 4320,
              marginTop: 300,
              marginRight: 300,
              marginBottom: 300,
              marginLeft: 300,
            },
          },
        ],
      },
      notes: {
        contents: [
          ...FOOTNOTES.map((id) => ({ id, noteKind: 'footnote', height: 0 })),
          ...ENDNOTES.map((id) => ({ id, noteKind: 'endnote', height: 0 })),
        ],
      },
      renderEnv: {},
    };
    const requirements = JSON.parse(
      session.layout_font_requirements_json(JSON.stringify(base))
    ) as Array<{
      key: string;
    }>;
    const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [fontId]]));
    const request = JSON.stringify({
      ...base,
      measurement: {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: true,
      },
    });
    const extras = JSON.stringify({ fontChains });
    const freshList = (): DisplayList => {
      const fresh = createEditSession(72);
      fresh.load(session.encode_state());
      fresh.load_note_separators(session.note_separators_state());
      fresh.register_measure_font(new Uint8Array(readFileSync(FONT)));
      fresh.layout_document_with_regions_json(request);
      return applyFrameDelta(null, decodeFrameDelta(fresh.build_display_list_frame(extras, 0)))
        .displayList;
    };

    session.layout_document_with_regions_json(request);
    // Each applier decodes its own copy: the owned one mutates what it retains.
    const full = session.build_display_list_frame(extras, 0);
    let owned = applyFrameDeltaOwned(null, decodeFrameDelta(full));
    let copied = applyFrameDelta(null, decodeFrameDelta(full));
    const areaKinds = owned.displayList.pages.map((page) =>
      (page.noteAreas ?? []).map((area) => area.kind).sort()
    );
    expect(areaKinds.at(-1)).toEqual(['endnote', 'footnote']);

    const updates: string[] = [];
    const engine: RustDisplayListQueryEngine = {
      hitTestRegionsJson: layoutWasm.hitTestRegionsJson,
      rangeRectsJson: layoutWasm.rangeRectsJson,
      hasDisplayListSession: layoutWasm.hasDisplayListSession,
      openDisplayList: layoutWasm.openDisplayList,
      closeDisplayList: layoutWasm.closeDisplayList,
      updateDisplayList: (handle, update) => {
        layoutWasm.updateDisplayList(handle, update);
        updates.push(update);
      },
      hasDisplayListUpdate: layoutWasm.hasDisplayListUpdate,
      hitTestRegionsByHandle: layoutWasm.hitTestRegionsByHandle,
      rangeRectsByHandle: layoutWasm.rangeRectsByHandle,
    };
    const firstQueries = createDisplayListQueries(owned.displayList, engine);
    firstQueries.prime();
    // parse every page into the store, so adoption has retained pages to shift
    firstQueries.rangeRects(0, Number.MAX_SAFE_INTEGER);
    const parsedUpdates = updates.length;

    const { paraId } = (JSON.parse(session.paragraphs('body')) as Array<{ paraId: string }>)[0]!;
    for (const text of ['xyz', 'ab']) {
      session.insert_text('body', paraId, 3, text, undefined, undefined);
      session.layout_document_with_regions_json(request);
      const frame = session.build_display_list_frame(extras, owned.frameEpoch);
      const delta = decodeFrameDelta(frame);
      const anchored = delta.operations.filter(
        (operation) => operation.kind === 'shift-positions' && operation.anchors.length > 0
      );
      expect(anchored.map((operation) => operation.pageIndex)).toEqual(
        areaKinds.flatMap((kinds, index) => (kinds.length > 0 ? [index] : []))
      );
      owned = applyFrameDeltaOwned(owned, delta);
      copied = applyFrameDelta(copied, decodeFrameDelta(frame));
    }

    const rebuilt = freshList();
    expect(owned.displayList).toEqual(rebuilt);
    expect(copied.displayList).toEqual(rebuilt);

    const lastPage = owned.displayList.pages.at(-1)!;
    expect(
      displayPageShiftsSince(lastPage, displayPageRevision(lastPage) - 2)!.map(
        (step) => step.anchors.length
      )
    ).toEqual([3, 3]);
    const secondQueries = createDisplayListQueries(owned.displayList, engine, firstQueries);
    secondQueries.prime();
    expect(updates).toHaveLength(parsedUpdates + 1);
    const update = JSON.parse(updates.at(-1)!) as {
      replace: Array<[number, unknown]>;
      shift: Array<[number, number, unknown[][], unknown[][]?]>;
    };
    expect(update.replace.map(([index]) => index)).toEqual([0]);
    const anchoredEntries = update.shift.filter((entry) => entry.length === 4);
    expect(anchoredEntries.map(([index]) => index)).toEqual(
      areaKinds.flatMap((kinds, index) => (kinds.length > 0 ? [index] : []))
    );
    expect(
      anchoredEntries.every(([, , runs, anchors]) => runs.length === 2 && anchors!.length === 2)
    ).toBe(true);

    const freshQueries = createDisplayListQueries(rebuilt, engine);
    for (const [from, to] of [
      [1, 40],
      [400, 900],
      [1300, 1500],
      [2500, 2900],
    ]) {
      expect(secondQueries.rangeRects(from!, to!)).toEqual(freshQueries.rangeRects(from!, to!));
    }
    for (const id of FOOTNOTES) {
      expect(secondQueries.noteRangeRects('footnote', id, 1, 4)).toEqual(
        freshQueries.noteRangeRects('footnote', id, 1, 4)
      );
    }
    for (const queries of [firstQueries, secondQueries, freshQueries]) queries.dispose();
  });
});

type WireValue = null | boolean | number | string | WireValue[] | WireObject;
interface WireObject {
  entries: [string, WireValue][];
}

/** One full frame holding a single page whose payload is `page`. */
function encodeSinglePageFrame(page: WireObject): Uint8Array {
  const strings: string[] = [];
  const stringId = (value: string): number => {
    const index = strings.indexOf(value);
    if (index >= 0) return index;
    strings.push(value);
    return strings.length - 1;
  };
  const u32 = (out: number[], value: number): void => {
    for (let shift = 0; shift < 32; shift += 8) out.push((value >>> shift) & 0xff);
  };
  const encode = (out: number[], value: WireValue): void => {
    if (value === null) out.push(0);
    else if (typeof value === 'boolean') out.push(value ? 2 : 1);
    else if (typeof value === 'number') {
      out.push(3);
      const bytes = new Uint8Array(8);
      new DataView(bytes.buffer).setBigInt64(0, BigInt(value), true);
      out.push(...bytes);
    } else if (typeof value === 'string') {
      out.push(6);
      u32(out, stringId(value));
    } else {
      const body: number[] = [];
      const items = Array.isArray(value) ? value : value.entries;
      for (const item of items) {
        if (Array.isArray(value)) {
          encode(body, item as WireValue);
        } else {
          const [key, entry] = item as [string, WireValue];
          u32(body, stringId(key));
          encode(body, entry);
        }
      }
      out.push(Array.isArray(value) ? 7 : 8);
      u32(out, body.length);
      u32(out, items.length);
      out.push(...body);
    }
  };
  const payload: number[] = [];
  encode(payload, page);

  const table: number[] = [];
  u32(table, strings.length);
  for (const value of strings) {
    const bytes = new TextEncoder().encode(value);
    u32(table, bytes.length);
    table.push(...bytes);
  }
  const stringsOffset = 80 + 48;
  const dataOffset = Math.ceil((stringsOffset + table.length) / 8) * 8;
  const total = dataOffset + payload.length;
  const bytes = new Uint8Array(total);
  const view = new DataView(bytes.buffer);
  bytes.set([0x46, 0x44, 0x56, 0x31], 0);
  view.setUint16(4, FRAME_DELTA_VERSION, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, total, true);
  view.setUint32(12, 1, true);
  view.setBigUint64(16, 1n, true);
  view.setBigUint64(24, 1n, true);
  view.setBigUint64(32, 1n, true);
  view.setUint32(48, 1, true);
  view.setUint32(52, 1, true);
  view.setUint32(56, 80, true);
  view.setUint32(60, stringsOffset, true);
  view.setUint32(64, table.length, true);
  view.setUint32(68, dataOffset, true);
  bytes[80] = 1;
  view.setBigUint64(88, 1n, true);
  view.setUint32(80 + 28, dataOffset, true);
  view.setUint32(80 + 32, dataOffset, true);
  view.setUint32(80 + 36, payload.length, true);
  bytes.set(table, stringsOffset);
  bytes.set(payload, dataOffset);
  return bytes;
}

describe('FrameDelta typed values', () => {
  const page = (meta: [string, WireValue][]): WireObject => ({
    entries: [
      ['pageIndex', 0],
      ['width', 10],
      ['height', 20],
      ['primitives', []],
      ['meta', { entries: meta }],
    ],
  });

  it('accepts a missing or zero watermark prefix on legacy pages', () => {
    for (const count of [undefined, 0]) {
      const payload = page([]);
      if (count !== undefined) payload.entries.push(['watermarkPrimitiveCount', count]);
      const decoded = decodeFrameDelta(encodeSinglePageFrame(payload));
      const upsert = decoded.operations[0];
      if (upsert?.kind !== 'upsert') throw new Error('expected an upsert');
      expect(upsert.page.watermarkPrimitiveCount).toBe(count);
    }
  });

  it('rejects an invalid watermark prefix', () => {
    for (const count of [-1, 1, '1']) {
      const payload = page([]);
      payload.entries.push(['watermarkPrimitiveCount', count]);
      expect(() => decodeFrameDelta(encodeSinglePageFrame(payload))).toThrow(
        'watermark primitive count is invalid'
      );
    }
  });

  it('decodes a __proto__ key as an own property without touching the prototype', () => {
    const decoded = decodeFrameDelta(
      encodeSinglePageFrame(page([['__proto__', { entries: [['polluted', true]] }], ['kept', 'yes']]))
    );
    const upsert = decoded.operations[0];
    if (upsert?.kind !== 'upsert') throw new Error('expected an upsert');
    const meta = (upsert.page as unknown as { meta: Record<string, unknown> }).meta;
    expect(Object.getPrototypeOf(meta)).toBe(Object.prototype);
    expect(Object.keys(meta)).toEqual(['__proto__', 'kept']);
    expect(Object.getOwnPropertyDescriptor(meta, '__proto__')?.value).toEqual({ polluted: true });
    expect((meta as { polluted?: unknown }).polluted).toBeUndefined();
    expect(upsert.page.width).toBe(10);
  });

  it('keeps a key that Object.prototype defines as an own property', () => {
    const decoded = decodeFrameDelta(
      encodeSinglePageFrame(page([['toString', 'wire'], ['constructor', 1]]))
    );
    const upsert = decoded.operations[0];
    if (upsert?.kind !== 'upsert') throw new Error('expected an upsert');
    const meta = (upsert.page as unknown as { meta: Record<string, unknown> }).meta;
    expect(Object.getOwnPropertyDescriptor(meta, 'toString')?.value).toBe('wire');
    expect(Object.getOwnPropertyDescriptor(meta, 'constructor')?.value).toBe(1);
    expect(() =>
      decodeFrameDelta(encodeSinglePageFrame(page([['toString', 1], ['toString', 2]])))
    ).toThrow('duplicate object key toString');
  });

  it('rejects a duplicate object key', () => {
    expect(() =>
      decodeFrameDelta(encodeSinglePageFrame(page([['kept', 1], ['kept', 2]])))
    ).toThrow('duplicate object key kept');
  });
});
