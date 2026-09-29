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

const PRESENT_ONLY = 1 << 7;

/** The first shift run's mask in `frame`, rewritten, must fail to decode when invalid. */
function expectStrictShiftMasks(frame: Uint8Array): void {
  const view = new DataView(frame.buffer, frame.byteOffset, frame.byteLength);
  const operations = view.getUint32(52, true);
  let maskOffset = -1;
  for (let index = 0; index < operations && maskOffset < 0; index++) {
    const record = 80 + index * 48;
    if (frame[record] === 5) maskOffset = view.getUint32(record + 32, true) + 16;
  }
  expect(maskOffset).toBeGreaterThan(0);
  for (const mask of [PRESENT_ONLY, (frame[maskOffset]! & 0x1f) | 0x40]) {
    const patched = frame.slice();
    patched[maskOffset] = mask;
    expect(() => decodeFrameDelta(patched)).toThrow('position shift run is invalid');
  }
}

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
    firstQueries.prime();

    session.insert_text('body', paraId, 5, 'x', undefined, undefined);
    const second = envelopeFor();
    const deltaFrame = session.build_display_list_frame(second, retained.frameEpoch);
    expectStrictShiftMasks(deltaFrame);
    const next = applyFrameDeltaOwned(retained, decodeFrameDelta(deltaFrame));

    // trailing pages absorb the insert as an in-place position shift with a
    // recorded, replayable run log
    expect(next.displayList.pages.at(-1)).toBe(trailingPage);
    expect(displayPageRevision(trailingPage)).toBe(revisionBefore + 1);
    const runLists = displayPageShiftsSince(trailingPage, revisionBefore);
    expect(runLists).not.toBeNull();
    expect(runLists!.length).toBe(1);
    // the whole page moves as one run over the fields each primitive has
    expect(runLists![0]).toHaveLength(1);
    expect(runLists![0][0]!.changedMask & PRESENT_ONLY).toBe(PRESENT_ONLY);
    expect(displayPageShiftsSince(trailingPage, revisionBefore + 1)).toEqual([]);

    // handle adoption ships those shifts as compact ops instead of replacing
    // the page's serialized payload
    const secondQueries = createDisplayListQueries(next.displayList, engine, firstQueries);
    secondQueries.prime();
    expect(updates.length).toBe(1);
    const update = JSON.parse(updates[0]!) as {
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
        operations: [{ kind: 'shift-positions', pageIndex: 0, pageId: 1n, fingerprint: 2n, runs }],
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
