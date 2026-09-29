import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { createEditSession, preloadEditWasm } from '../../wasm/edit';
import {
  applyFrameDelta,
  applyFrameDeltaOwned,
  decodeFrameDelta,
  displayPageRevision,
  displayPageShiftsSince,
  FRAME_DELTA_VERSION,
  type DecodedFrameDelta,
  type RetainedFrame,
} from './frameDelta';
import { createDisplayListQueries } from './displayListQueries';
import type { RustDisplayListQueryEngine } from './rustDisplayList';
import type { DisplayList, DisplayPage } from './displayList';

const WASM = resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

describe('FrameDelta wire round-trip', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('decodes wasm-encoded full and delta frames to the equivalent JSON list', () => {
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
      });
    };

    const first = envelopeFor();
    const jsonList = JSON.parse(session.build_display_list_json(first)) as DisplayList;
    const fullFrame = session.build_display_list_frame(first, 0);
    const retained = applyFrameDelta(null, decodeFrameDelta(fullFrame));
    expect(retained.displayList).toEqual(jsonList);

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

  it('records owned position shifts and ships them as query-store shift ops', () => {
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
    const next = applyFrameDeltaOwned(
      retained,
      decodeFrameDelta(session.build_display_list_frame(second, retained.frameEpoch))
    );

    // trailing pages absorb the insert as an in-place position shift with a
    // recorded, replayable run log
    expect(next.displayList.pages.at(-1)).toBe(trailingPage);
    expect(displayPageRevision(trailingPage)).toBe(revisionBefore + 1);
    const runLists = displayPageShiftsSince(trailingPage, revisionBefore);
    expect(runLists).not.toBeNull();
    expect(runLists!.length).toBe(1);
    expect(runLists![0].length).toBeGreaterThan(0);
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
        },
      ],
      bytes: new Uint8Array(),
    };

    expect(() => applyFrameDeltaOwned(previous, delta)).toThrow('requires retained docStart');
    expect(displayPageRevision(page)).toBe(0);
    expect(displayPageShiftsSince(page, 0)).toEqual([]);
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
