import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { runInNewContext } from 'node:vm';
import {
  createSessionClient, createSessionHost, isHostMessage, SESSION_SUPERSEDED, SessionFailure,
  type MethodPolicy, type SessionTransport,
} from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { XlsxEditRequest, XlsxRangeTarget, XlsxReadResult } from '../edits';
import { initWasm, openWorkbook, type Viewport, type WorkbookHandle } from '../wasm/loader';
import { createWorkbookSession, type WorkbookSession } from './client';
import { createWorkbookSessionHost } from './host';
import {
  WORKBOOK_SESSION_METHODS,
  WORKBOOK_SESSION_POLICIES,
  type WorkbookFrame,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
  type WorkbookSheetView,
} from './methods';

let fixture: Uint8Array;
let chartFixture: Uint8Array;
let wasmBytes: Uint8Array<ArrayBuffer>;
const viewport: Viewport = { x: 0, y: 0, width: 800, height: 800 };

beforeAll(async () => {
  const [wasm, xlsx, charts] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/xlsx_wasm_bg.wasm')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/charts.xlsx')),
  ]);
  wasmBytes = new Uint8Array(wasm);
  await initWasm(wasmBytes);
  fixture = new Uint8Array(xlsx);
  chartFixture = new Uint8Array(charts);
});

function target(a1: string, sheetId = 'sheet:0'): XlsxRangeTarget {
  return { sheetId, range: { kind: 'a1', a1 } };
}

function read(result: XlsxReadResult): Extract<XlsxReadResult, { ok: true }> {
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

function content(result: XlsxReadResult): Omit<Extract<XlsxReadResult, { ok: true }>, 'version'> {
  const { version: _, ...value } = read(result);
  return value;
}

function editBatch(expectVersion: string): XlsxEditRequest {
  return {
    expectVersion,
    steps: [
      { op: 'setCellInputs', target: target('B3'), inputs: [['1000']] },
      { op: 'patchStyle', target: target('B3'), patch: { bold: true } },
      { op: 'setNumberFormat', target: target('C3'), format: 'percent' },
    ],
  };
}

function batchRequests(transport: SessionTransport): SessionTransport {
  return { ...transport, listen(listener) {
    let messages: unknown[] = [];
    let timer: ReturnType<typeof setTimeout> | undefined;
    const off = transport.listen((message) => {
      messages.push(message);
      if (timer !== undefined) return;
      timer = setTimeout(() => {
        timer = undefined;
        const batch = messages;
        messages = [];
        for (const queued of batch) listener(queued);
      }, 0);
    });
    return () => {
      off();
      if (timer !== undefined) clearTimeout(timer);
      messages = [];
    };
  } };
}

async function session(
  bytes = fixture, wrapHost?: (transport: SessionTransport) => SessionTransport
): Promise<WorkbookSession> {
  const pair = createInProcessPair();
  createWorkbookSessionHost(wrapHost ? wrapHost(pair.host) : pair.host);
  return createWorkbookSession(bytes, {}, pair.client);
}

async function matchingFrame(worker: WorkbookSession, main: WorkbookHandle): Promise<WorkbookFrame> {
  const frame = await worker.call.frame(viewport);
  expect(frame.displayList).toEqual(main.displayList(viewport));
  expect(frame.version).toBe(await worker.call.version());
  expect(frame.sheet).toBe(main.sheetInfo().activeSheet);
  expect(frame.viewport).toEqual(viewport);
  return frame;
}

async function matchingSheetView(
  worker: WorkbookSession, main: WorkbookHandle, sheet: number
): Promise<WorkbookSheetView> {
  main.setActiveSheet(sheet);
  const { contentWidth, contentHeight, frozenRows, frozenCols, initialScrollX, initialScrollY } =
    main.sheetInfo();
  const grid = main.displayList(viewport).grid;
  if (!grid) throw new Error('Missing frame grid');
  const view = await worker.call.sheetView(sheet);
  expect(view).toEqual({
    sheet, version: await worker.call.version(), contentWidth, contentHeight, frozenRows, frozenCols,
    initialScrollX, initialScrollY,
    frozenWidth: frozenCols === 0 ? 0 : grid.colOffsets[frozenCols],
    frozenHeight: frozenRows === 0 ? 0 : grid.rowOffsets[frozenRows],
  });
  return view;
}

describe('workbook sessions', () => {
  test('declares ordered methods and marks edits as user input', () => {
    expect(Object.keys(WORKBOOK_SESSION_METHODS)).toEqual(Object.keys(WORKBOOK_SESSION_POLICIES));
    expect(WORKBOOK_SESSION_POLICIES.applyEdits).toEqual({
      lane: 'input', mutates: true, userInput: true,
    });
    expect(WORKBOOK_SESSION_POLICIES.frame).toEqual({
      lane: 'interactive', reframes: true, key: 'frame', replaceableBy: 'frame',
    });
    for (const policy of Object.values(WORKBOOK_SESSION_POLICIES)) {
      expect(typeof policy).toBe('object');
      expect((policy as MethodPolicy).reorderable).not.toBe(true);
    }
  });

  for (const name of ['sample.xlsx', 'charts.xlsx'] as const) {
    test(`matches ${name} view reads without changing active sheet, state or saved bytes`, async () => {
      const bytes = name === 'sample.xlsx' ? fixture : chartFixture;
      const main = openWorkbook(bytes);
      let worker: WorkbookSession | undefined;
      try {
        worker = await session(bytes);
        const state = worker.state;
        const version = await worker.call.version();
        const saved = await worker.save();
        const changes: WorkbookSessionEvents['changed'][] = [];
        worker.on('changed', (change) => { changes.push(change); });
        for (const sheet of state.sheets) {
          const view = await matchingSheetView(worker, main, sheet.index);
          if (name === 'charts.xlsx' && sheet.index === 0) {
            expect(view).toMatchObject({ frozenRows: 1, frozenCols: 1 });
            expect(view.frozenWidth).toBeGreaterThan(0);
            expect(view.frozenHeight).toBeGreaterThan(0);
          }
          for (const [row, col] of [[0, 0], [2, 1], [30, 10], [1048575, 16383]]) {
            expect(await worker.call.cellGeometry(sheet.index, row, col)).toEqual({
              sheet: sheet.index, version, rect: main.cellRect(sheet.index, row, col),
              scrollPosition: main.cellPosition(sheet.index, row, col),
            });
          }
          const inputs = await worker.call.cellInputs(sheet.index, 'A1:D5');
          expect(inputs).toEqual({
            sheet: sheet.index, version, cells: main.rangeCells(sheet.index, 'A1:D5'),
          });
          expect(inputs.cells[0][0]).toEqual(main.cell(sheet.index, 0, 0));
          for (const window of [viewport, { ...viewport, x: 1500, y: 1500 }]) {
            const frame = await worker.call.frame(window, { sheet: sheet.index });
            expect(frame.displayList).toEqual(main.displayList(window));
            const expected = name === 'sample.xlsx' && sheet.index === 0 && window === viewport
              ? [{
                start: { row: 0, col: 0 }, end: { row: 0, col: 3 },
              }] : [];
            expect(frame.mergedRanges).toEqual(expected);
          }
        }
        main.setActiveSheet(state.activeSheet);
        await matchingFrame(worker, main);
        expect(worker.state).toEqual(state);
        expect(await worker.call.version()).toBe(version);
        expect(await worker.save()).toEqual(saved);
        expect(changes).toEqual([]);
      } finally {
        main.dispose();
        await worker?.dispose();
      }
    });
  }

  test('view metadata, geometry and editable input follow edits and their version', async () => {
    const main = openWorkbook(fixture);
    let worker: WorkbookSession | undefined;
    try {
      worker = await session();
      const before = await matchingSheetView(worker, main, 0);
      const geometry = await worker.call.cellGeometry(0, 69, 0);
      const batch = (expectVersion: string): XlsxEditRequest => ({
        expectVersion,
        steps: [
          { op: 'setCellInputs', target: target('A70'), inputs: [['updated']] },
          { op: 'patchStyle', target: target('A70'), patch: { fontSize: 40 } },
          { op: 'setCellInputs', target: target('B8:F8'), inputs: [[
            '=1+2', "'=literal", "'123", 'TRUE', '12.5',
          ]] },
        ],
      });
      const applied = await worker.call.applyEdits(batch(before.version));
      expect(applied).toMatchObject({ ok: true, applied: true });
      expect(main.applyEdits(batch(main.version()))).toMatchObject({ ok: true, applied: true });
      const after = await matchingSheetView(worker, main, 0);
      expect(after.version).toBe(applied.version);
      expect(after.version).not.toBe(before.version);
      expect(after.contentHeight).toBeGreaterThan(before.contentHeight);
      const moved = await worker.call.cellGeometry(0, 69, 0);
      expect(moved.version).toBe(applied.version);
      expect(moved.rect).toEqual(main.cellRect(0, 69, 0));
      expect(moved.rect.h).toBeGreaterThan(geometry.rect.h);
      expect(moved.scrollPosition).toEqual(main.cellPosition(0, 69, 0));
      const inputs = await worker.call.cellInputs(0, 'B8:F8');
      expect(inputs.version).toBe(applied.version);
      expect(inputs.cells).toEqual(main.rangeCells(0, 'B8:F8'));
      for (let col = 1; col <= 5; col += 1) {
        expect(inputs.cells[0][col - 1]).toEqual(main.cell(0, 7, col));
      }
      expect(inputs.cells[0][0]).toMatchObject({ input: '=1+2', isFormula: true });
      const frame = await matchingFrame(worker, main);
      expect(frame.mergedRanges).toEqual([
        { start: { row: 0, col: 0 }, end: { row: 0, col: 3 } },
      ]);
      expect(frame.version).toBe(applied.version);
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  test('frame merges cross every viewport edge and exclude frozen gaps', async () => {
    const main = openWorkbook(chartFixture);
    let worker: WorkbookSession | undefined;
    try {
      const expected = [
        { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
        { start: { row: 8, col: 11 }, end: { row: 10, col: 11 } },
        { start: { row: 11, col: 8 }, end: { row: 11, col: 10 } },
        { start: { row: 12, col: 13 }, end: { row: 12, col: 15 } },
        { start: { row: 13, col: 12 }, end: { row: 15, col: 12 } },
        { start: { row: 8, col: 8 }, end: { row: 10, col: 10 } },
        { start: { row: 1, col: 1 }, end: { row: 1, col: 10 } },
        { start: { row: 1, col: 0 }, end: { row: 10, col: 0 } },
        { start: { row: 11, col: 11 }, end: { row: 11, col: 12 } },
      ];
      const excluded = [
        { start: { row: 3, col: 3 }, end: { row: 4, col: 4 } },
        { start: { row: 3, col: 11 }, end: { row: 4, col: 11 } },
        { start: { row: 12, col: 3 }, end: { row: 12, col: 4 } },
        { start: { row: 16, col: 16 }, end: { row: 17, col: 17 } },
      ];
      expect(main.applyOps([
        { type: 'setFreezePane', sheet: 0, pane: {
          rows: 2, cols: 2, top_left: { row: 10, col: 10 },
        } },
        ...[...expected, ...excluded].map((range) => ({ type: 'mergeCells', sheet: 0, range })),
      ]).applied).toBe(true);
      worker = await session(main.save());
      const position = main.cellPosition(0, 10, 10);
      const frozen = main.cellRect(0, 1, 1);
      const first = main.cellRect(0, 10, 10);
      const end = main.cellRect(0, 14, 14);
      const window = {
        x: position.x + first.w / 4,
        y: position.y + first.h / 4,
        width: frozen.x + frozen.w + end.x - first.x - first.w / 2,
        height: frozen.y + frozen.h + end.y - first.y - first.h / 2,
      };
      const frame = await worker.call.frame(window, { sheet: 0 });
      expect(frame.displayList).toEqual(main.displayList(window));
      expect(frame.displayList.grid?.rowIndices).toEqual([0, 1, 10, 11, 12, 13]);
      expect(frame.displayList.grid?.colIndices).toEqual([0, 1, 10, 11, 12, 13]);
      expect(frame.mergedRanges).toEqual(expected);
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  test('matches main-thread projections, reads, edits and saved bytes', async () => {
    const main = openWorkbook(fixture);
    let worker: WorkbookSession | undefined;
    try {
      worker = await session();
      const info = main.sheetInfo();
      const summaries = info.sheetIds.map((id, index) => ({
        id, index, name: info.sheetNames[index],
      }));
      expect(worker.state).toEqual({
        format: 'xlsx', stage: 'ready', version: 0, dirty: false,
        sheets: summaries, activeSheet: info.activeSheet,
      });
      expect(await worker.call.sheets()).toEqual(summaries);
      expect(await worker.call.calculationStatus()).toEqual(main.calculationStatus());
      expect(await worker.save()).toEqual(main.save());
      const query = { ranges: [target('B3:D3')] };
      const mainBefore = read(main.readCells(query));
      const workerBefore = read(await worker.call.readCells(query));
      expect(await worker.call.version()).toBe(workerBefore.version);
      expect(mainBefore.version).toBe(main.version());
      expect(workerBefore.version).not.toBe(mainBefore.version);
      expect(content(workerBefore)).toEqual(content(mainBefore));
      const request = editBatch(workerBefore.version);
      expect(await worker.call.validateEdits(request)).toMatchObject({ ok: true, wouldApply: true });
      expect(worker.state.dirty).toBe(false);
      const applied = await worker.call.applyEdits(request);
      const mainApplied = main.applyEdits({ ...request, expectVersion: mainBefore.version });
      if (!applied.ok || !mainApplied.ok) throw new Error('Parity batch was refused');
      expect(applied.applied).toBe(true);
      expect(applied.baseVersion).toBe(workerBefore.version);
      expect(mainApplied.baseVersion).toBe(mainBefore.version);
      const { baseVersion: workerBase, version: workerVersion, ...workerReceipt } = applied;
      const { baseVersion: mainBase, version: mainVersion, ...mainReceipt } = mainApplied;
      expect(workerBase).not.toBe(mainBase);
      expect(workerVersion).not.toBe(mainVersion);
      expect(workerReceipt).toEqual(mainReceipt);
      expect(await worker.call.version()).toBe(applied.version);
      expect(main.version()).toBe(mainApplied.version);
      expect(applied.version).not.toBe(workerBefore.version);
      expect(mainApplied.version).not.toBe(mainBefore.version);
      const workerAfter = read(await worker.call.readCells(query));
      const mainAfter = read(main.readCells(query));
      expect(workerAfter.version).toBe(applied.version);
      expect(mainAfter.version).toBe(mainApplied.version);
      expect(content(workerAfter)).toEqual(content(mainAfter));
      const found = await worker.call.findText({ text: 'Quarterly' });
      const mainFound = main.findText({ text: 'Quarterly' });
      expect(found.version).toBe(applied.version);
      expect(mainFound.version).toBe(mainApplied.version);
      const { version: foundVersion, ...matches } = found;
      const { version: mainFoundVersion, ...mainMatches } = mainFound;
      expect(foundVersion).not.toBe(mainFoundVersion);
      expect(matches).toEqual(mainMatches);
      expect(await worker.call.calculationStatus()).toEqual(main.calculationStatus());
      expect(await worker.save()).toEqual(main.save());
      expect<Uint8Array>(new Uint8Array(await worker.call.save())).toEqual(main.save());
      expect(worker.state).toMatchObject({ version: 1, dirty: true });
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  for (const name of ['sample.xlsx', 'charts.xlsx'] as const) {
    test(`matches ${name} frames at open and after edits with increasing epochs`, async () => {
      const bytes = name === 'sample.xlsx' ? fixture : chartFixture;
      const main = openWorkbook(bytes);
      let worker: WorkbookSession | undefined;
      try {
        worker = await session(bytes);
        const before = await matchingFrame(worker, main);
        expect(before.epoch).toBe(1);
        const applied = await worker.call.applyEdits(editBatch(before.version));
        const mainApplied = main.applyEdits(editBatch(main.version()));
        expect(applied).toMatchObject({ ok: true, applied: true });
        expect(mainApplied).toMatchObject({ ok: true, applied: true });
        const after = await matchingFrame(worker, main);
        expect(after.epoch).toBe(before.epoch + 1);
        expect(after.version).not.toBe(before.version);
        expect(after.displayList).not.toEqual(before.displayList);
        if (name === 'charts.xlsx') {
          for (const frame of [before, after]) {
            expect(frame.displayList.charts?.length).toBe(4);
            expect(frame.displayList.commands.some((command) => command.op === 'path')).toBe(true);
            expect(frame.displayList.commands.some((command) =>
              command.op === 'text' && command.chart && command.text === 'Revenue trend'
            )).toBe(true);
          }
        }
      } finally {
        main.dispose();
        await worker?.dispose();
      }
    });
  }

  test('frames a second sheet without changing state, events, version or saved bytes', async () => {
    const main = openWorkbook(fixture);
    const workers: WorkbookSession[] = [];
    try {
      workers.push(await session());
      workers.push(await session());
      const [worker, untouched] = workers;
      const state = worker.state;
      const version = await worker.call.version();
      const changes: WorkbookSessionEvents['changed'][] = [];
      worker.on('changed', (change) => { changes.push(change); });
      expect(state.sheets.length).toBeGreaterThan(1);
      main.setActiveSheet(1);
      const frame = await worker.call.frame(viewport, { sheet: 1 });
      expect(frame.sheet).toBe(1);
      expect(frame.displayList).toEqual(main.displayList(viewport));
      expect(frame.version).toBe(version);
      main.setActiveSheet(state.activeSheet);
      await matchingFrame(worker, main);
      expect(worker.state).toEqual(state);
      expect(worker.state).toEqual(untouched.state);
      expect(await worker.call.sheets()).toEqual(await untouched.call.sheets());
      expect(changes).toEqual([]);
      expect(await worker.save()).toEqual(await untouched.save());
    } finally {
      main.dispose();
      for (const worker of workers) await worker.dispose();
    }
  });

  test('supersedes the middle of three queued frames and remains usable', async () => {
    const worker = await session(fixture, batchRequests);
    try {
      const lastViewport = { ...viewport, x: 128, y: 64 };
      const results = await Promise.allSettled([
        worker.call.frame(viewport),
        worker.call.frame({ ...viewport, x: 64 }),
        worker.call.frame(lastViewport),
      ]);
      for (const result of results.slice(0, 2)) {
        if (result.status !== 'rejected') throw new Error('Queued frame was not superseded');
        expect(result.reason).toMatchObject({ name: SESSION_SUPERSEDED });
      }
      const last = results[2];
      if (last.status !== 'fulfilled') throw last.reason;
      expect(last.value.viewport).toEqual(lastViewport);
      expect(last.value.epoch).toBe(1);
      expect(worker.failure).toBeUndefined();
      expect((await worker.call.frame(viewport)).epoch).toBe(2);
    } finally { await worker.dispose(); }
  });

  test('keeps edits between queued frames ahead of the later frame', async () => {
    const main = openWorkbook(fixture);
    let worker: WorkbookSession | undefined;
    try {
      worker = await session(fixture, batchRequests);
      const before = await matchingFrame(worker, main);
      const replaced = worker.call.frame({ ...viewport, x: 64 }).catch((error: unknown) => error);
      const editing = worker.call.applyEdits(editBatch(before.version));
      const framing = worker.call.frame(viewport);
      const [dropped, applied, after] = await Promise.all([replaced, editing, framing]);
      expect(dropped).toMatchObject({ name: SESSION_SUPERSEDED });
      if (!applied.ok) throw new Error(applied.failure.message);
      expect(applied.applied).toBe(true);
      expect(main.applyEdits(editBatch(main.version()))).toMatchObject({ ok: true, applied: true });
      expect(after.displayList).toEqual(main.displayList(viewport));
      expect(after.version).toBe(applied.version);
      expect(after.version).not.toBe(before.version);
      expect(after.epoch).toBe(before.epoch + 1);
      expect(worker.state).toMatchObject({ version: 1, dirty: true });
      expect(worker.failure).toBeUndefined();
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  test('transfers frame JSON as an ArrayBuffer without command arrays', async () => {
    const buffers: ArrayBuffer[] = [];
    const worker = await session(fixture, (transport) => ({ ...transport, post(message, transfer) {
      if (isHostMessage(message) && message.kind === 'reply' && message.ok &&
        message.value !== null && typeof message.value === 'object' && 'epoch' in message.value) {
        const wire = message.value as unknown as { displayList: ArrayBuffer };
        expect(wire.displayList).toBeInstanceOf(ArrayBuffer);
        expect(transfer).toEqual([wire.displayList]);
        expect(JSON.stringify(message)).not.toContain('"commands"');
        expect(wire.displayList.byteLength).toBeGreaterThan(0);
        buffers.push(wire.displayList);
      }
      transport.post(message, transfer);
    } }));
    const main = openWorkbook(fixture);
    try {
      await matchingFrame(worker, main);
      expect(buffers).toHaveLength(1);
      expect(buffers[0].byteLength).toBe(0);
    } finally {
      main.dispose();
      await worker.dispose();
    }
  });

  test('rejects invalid sheet indices on direct handle view reads', () => {
    const main = openWorkbook(fixture);
    try {
      for (const sheet of [-1, 0.5, main.sheetCount(), 2 ** 32, NaN]) {
        expect(() => main.visibleMergedRanges(sheet, viewport)).toThrow(RangeError);
        expect(() => main.sheetInfoFor(sheet)).toThrow(RangeError);
      }
    } finally { main.dispose(); }
  });

  test('rejects invalid viewports and sheet indices without changing the session', async () => {
    const worker = await session();
    try {
      const state = worker.state;
      const version = await worker.call.version();
      const saved = await worker.save();
      for (const invalid of [
        { ...viewport, x: NaN }, { ...viewport, height: Infinity },
        { ...viewport, width: 0 }, { ...viewport, width: 1e9, height: 1e9 },
      ]) await expect(worker.call.frame(invalid)).rejects.toThrow();
      for (const sheet of [-1, 0.5, state.sheets.length, 2 ** 32, NaN]) {
        await expect(worker.call.frame(viewport, { sheet })).rejects.toMatchObject({
          name: 'RangeError',
        });
        await expect(worker.call.sheetView(sheet)).rejects.toMatchObject({ name: 'RangeError' });
        await expect(worker.call.cellGeometry(sheet, 0, 0)).rejects.toMatchObject({ name: 'RangeError' });
        await expect(worker.call.cellInputs(sheet, 'A1')).rejects.toMatchObject({ name: 'RangeError' });
      }
      for (const [row, col] of [[-1, 0], [0.5, 0], [1048576, 0], [0, 16384], [NaN, 0]]) {
        await expect(worker.call.cellGeometry(0, row, col)).rejects.toThrow();
      }
      await expect(worker.call.cellInputs(0, 'not a range')).rejects.toThrow();
      expect(worker.state).toEqual(state);
      expect(worker.failure).toBeUndefined();
      const frame = await worker.call.frame(viewport);
      expect(frame.epoch).toBe(1);
      expect(frame.version).toBe(version);
      expect(await worker.save()).toEqual(saved);
    } finally { await worker.dispose(); }
  });

  test('returns refusals as data and orders edits, reads, saves and changed events', async () => {
    const worker = await session();
    const changes: WorkbookSessionEvents['changed'][] = [];
    const off = worker.on('changed', (change) => { changes.push(change); });
    try {
      const initial = await worker.call.version();
      const query = { ranges: [target('B3')] };
      const request: XlsxEditRequest = {
        expectVersion: initial,
        steps: [{ op: 'setCellInputs', target: target('B3'), inputs: [['1000']] }],
      };
      expect(await worker.call.applyEdits({ ...request, expectVersion: 'stale' }))
        .toMatchObject({ ok: false, version: initial, failure: { code: 'stale-version' } });
      expect(await worker.call.validateEdits({ ...request, expectVersion: 'stale' }))
        .toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      expect(await worker.call.applyEdits({
        expectVersion: initial,
        steps: [{ op: 'setCellInputs', target: target('B3'), inputs: [['100']] }],
      })).toMatchObject({ ok: true, applied: false, version: initial });
      expect(worker.state).toMatchObject({ version: 0, dirty: false });
      expect(changes).toEqual([]);
      const first = await worker.call.applyEdits(request);
      if (!first.ok) throw new Error(first.failure.message);
      expect(first.applied).toBe(true);
      expect(await worker.call.applyEdits(request)).toMatchObject({
        ok: false, version: first.version, failure: { code: 'stale-version' },
      });
      const noOp = await worker.call.applyEdits({ ...request, expectVersion: first.version });
      expect(noOp).toMatchObject({ ok: true, applied: false, version: first.version });
      expect(changes).toEqual([{ version: 1, dirty: true }]);
      const beforeWrite = worker.call.readCells(query);
      const saveBefore = worker.save();
      const second = worker.call.applyEdits({
        expectVersion: first.version,
        steps: [{ op: 'setCellInputs', target: target('B3'), inputs: [['2000']] }],
      });
      const afterWrite = worker.call.readCells(query);
      const saveAfter = worker.save();
      const [before, savedBefore, applied, after, savedAfter] =
        await Promise.all([beforeWrite, saveBefore, second, afterWrite, saveAfter]);
      if (!applied.ok) throw new Error(applied.failure.message);
      expect(applied.applied).toBe(true);
      expect(read(before).version).toBe(first.version);
      expect(read(before).ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 1000 });
      expect(read(after).version).toBe(applied.version);
      expect(read(after).ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 2000 });
      expect(savedBefore).not.toEqual(savedAfter);
      expect(changes).toEqual([{ version: 1, dirty: true }, { version: 2, dirty: true }]);
      expect(worker.state).toMatchObject({ version: 2, dirty: true });
      expect(worker.failure).toBeUndefined();
      off();
      await worker.call.applyEdits({
        expectVersion: applied.version,
        steps: [{ op: 'setCellInputs', target: target('B3'), inputs: [['3000']] }],
      });
      expect(changes).toHaveLength(2);
    } finally {
      off();
      await worker.dispose();
    }
    await expect(worker.call.version()).rejects.toMatchObject({ code: 'disposed' });
    await expect(worker.call.frame(viewport)).rejects.toMatchObject({ code: 'disposed' });
    await expect(worker.save()).rejects.toMatchObject({ code: 'disposed' });
    await worker.dispose();
  });

  test('refuses calls before open, a second open, and calls after RPC disposal', async () => {
    const pair = createInProcessPair();
    let initializations = 0;
    createWorkbookSessionHost(pair.host, { initWasm: async (source) => {
      initializations += 1;
      await initWasm(source);
    } });
    const client = createSessionClient<WorkbookSessionMethods, {
      changed: WorkbookSessionEvents['changed'];
    }>(pair.client, { methods: WORKBOOK_SESSION_METHODS });
    try {
      const calls = [
        () => client.call.version(), () => client.call.readCells({ ranges: [] }),
        () => client.call.findText({ text: 'text' }),
        () => client.call.validateEdits({ expectVersion: 'stale', steps: [] }),
        () => client.call.applyEdits({ expectVersion: 'stale', steps: [] }),
        () => client.call.frame(viewport),
        () => client.call.sheetView(0), () => client.call.cellGeometry(0, 0, 0),
        () => client.call.cellInputs(0, 'A1'),
        () => client.call.sheets(), () => client.call.calculationStatus(),
        () => client.call.save(), () => client.call.dispose(),
      ];
      for (const call of calls) await expect(call()).rejects.toThrow('not open');
      expect(initializations).toBe(0);
      const bytes = new Uint8Array(fixture).buffer;
      await client.call.open(bytes);
      await expect(client.call.open(bytes)).rejects.toThrow('already open');
      expect(initializations).toBe(1);
      await client.call.dispose();
      for (const call of calls) await expect(call()).rejects.toThrow('disposed');
      await expect(client.call.open(bytes)).rejects.toThrow('disposed');
    } finally { await client.dispose(); }
  });

  test('opens a collaborative workbook with the requested client id', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const main = openWorkbook(fixture, { collaborative: true, clientId: 9701 });
    let worker: WorkbookSession | undefined;
    try {
      worker = await createWorkbookSession(fixture, {
        collaborative: true, clientId: 9701,
      }, pair.client);
      expect(worker.state.sheets).toEqual(main.sheetInfo().sheetIds.map((id, index) => ({
        id, index, name: main.sheetInfo().sheetNames[index],
      })));
      expect(content(await worker.call.readCells({ ranges: [] })))
        .toEqual(content(main.readCells({ ranges: [] })));
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  test('transfers owned copies of document and wasm buffers, including subviews and other realms', async () => {
    const main = openWorkbook(fixture);
    let saved: Uint8Array;
    try { saved = main.save(); } finally { main.dispose(); }
    for (const kind of ['buffer', 'view', 'foreign'] as const) {
      const asView = kind === 'view';
      const buffer: ArrayBuffer = kind === 'foreign'
        ? runInNewContext('new ArrayBuffer(size)', { size: fixture.byteLength })
        : new ArrayBuffer(fixture.byteLength + (asView ? 16 : 0));
      const wasm: ArrayBuffer = kind === 'foreign'
        ? runInNewContext('new ArrayBuffer(size)', { size: wasmBytes.byteLength })
        : new ArrayBuffer(wasmBytes.byteLength);
      if (kind === 'foreign') {
        expect(buffer instanceof ArrayBuffer).toBe(false);
        expect(wasm instanceof ArrayBuffer).toBe(false);
      }
      new Uint8Array(wasm).set(wasmBytes);
      const source = new Uint8Array(buffer);
      source.set(fixture, asView ? 8 : 0);
      const retainedSource = source.slice();
      const document = asView ? source.subarray(8, source.byteLength - 8) : source.buffer;
      const pair = createInProcessPair();
      createWorkbookSessionHost(pair.host, { initWasm: async (input) => {
        expect(input instanceof ArrayBuffer).toBe(true);
        expect(new Uint8Array(input as ArrayBuffer)).toEqual(wasmBytes);
        await initWasm(input);
      } });
      const transferred: ArrayBuffer[] = [];
      const transport: SessionTransport = { ...pair.client, post(message, transfer) {
        transferred.push(...(transfer ?? []) as ArrayBuffer[]);
        pair.client.post(message, transfer);
      } };
      const worker = await createWorkbookSession(document, { wasm }, transport);
      try {
        expect(transferred).toHaveLength(2);
        expect(transferred.every((value) => value.byteLength === 0)).toBe(true);
        expect(source.byteLength).toBe(fixture.byteLength + (asView ? 16 : 0));
        expect(source).toEqual(retainedSource);
        expect<Uint8Array>(new Uint8Array(document)).toEqual(fixture);
        expect(wasm.byteLength).toBe(wasmBytes.byteLength);
        expect<Uint8Array>(new Uint8Array(wasm)).toEqual(wasmBytes);
        expect(await worker.save()).toEqual(saved);
      } finally { await worker.dispose(); }
    }
  });

  test('sets the failed stage before notifying listeners of traps and crashes', async () => {
    for (const failure of [
      new WebAssembly.RuntimeError('unreachable'), new SessionFailure('crash', 'Host crashed'),
    ]) {
      const pair = createInProcessPair();
      createSessionHost<Pick<WorkbookSessionMethods, 'open' | 'version' | 'dispose'>, {}, null>(
        pair.host, {
          context: null,
          policies: {
            open: WORKBOOK_SESSION_POLICIES.open,
            version: WORKBOOK_SESSION_POLICIES.version,
            dispose: WORKBOOK_SESSION_POLICIES.dispose,
          },
          handlers: {
            open: () => ({
              format: 'xlsx', stage: 'ready', version: 0, dirty: false,
              sheets: [], activeSheet: 0,
            }),
            version: () => { throw failure; },
            dispose: () => {},
          },
        }
      );
      const worker = await createWorkbookSession(new Uint8Array(), {}, pair.client);
      const stages: WorkbookSession['state']['stage'][] = [];
      worker.onFailure(() => { stages.push(worker.state.stage); });
      try {
        expect(worker.state.stage).toBe('ready');
        await expect(worker.call.version()).rejects.toMatchObject({
          code: failure instanceof SessionFailure ? 'crash' : 'trap',
        });
        expect(stages).toEqual(['failed']);
        expect(worker.state.stage).toBe('failed');
      } finally { await worker.dispose(); }
    }
  });

  test('closes the transport when the caller buffers cannot be copied', async () => {
    if (typeof globalThis.structuredClone !== 'function') throw new Error('Missing structuredClone');
    const pair = createInProcessPair();
    const detached = new ArrayBuffer(8);
    globalThis.structuredClone(detached, { transfer: [detached] });
    await expect(createWorkbookSession(detached, {}, pair.client)).rejects.toThrow();
    expect(() => pair.host.post({})).toThrow();
  });

  test('closes the transport when the session client cannot attach', async () => {
    const pair = createInProcessPair();
    let closed = 0;
    const transport: SessionTransport = {
      post: (message, transfer) => pair.client.post(message, transfer),
      listen: (listener) => pair.client.listen(listener),
      onError() { throw new Error('attach'); },
      close() { closed += 1; pair.client.close(); },
    };
    await expect(createWorkbookSession(fixture, {}, transport)).rejects.toThrow('attach');
    expect(closed).toBe(1);
  });

  test('an aborted signal closes the transport of an open in flight', async () => {
    const pair = createInProcessPair();
    let closed = 0;
    const transport: SessionTransport = {
      post: (message, transfer) => pair.client.post(message, transfer),
      listen: (listener) => pair.client.listen(listener),
      onError: (listener) => pair.client.onError(listener),
      close() { closed += 1; pair.client.close(); },
    };
    const controller = new AbortController();
    const opening = createWorkbookSession(fixture, { signal: controller.signal }, transport);
    controller.abort();
    await expect(opening).rejects.toMatchObject({
      name: 'SessionFailure', code: 'disposed', message: 'Session was disposed',
    });
    expect(closed).toBeGreaterThan(0);
  });

  test('an abort after the open reply rejects before returning the session', async () => {
    const pair = createInProcessPair();
    const controller = new AbortController();
    let closed = 0;
    const transport: SessionTransport = {
      post: (message, transfer) => pair.client.post(message, transfer),
      listen: (listener) => pair.client.listen((message) => {
        listener(message);
        if (isHostMessage(message) && message.kind === 'reply' && message.ok) controller.abort();
      }),
      onError: (listener) => pair.client.onError(listener),
      close() { closed += 1; pair.client.close(); },
    };
    const opening = createWorkbookSession(fixture, { signal: controller.signal }, transport);
    pair.host.post({
      protocol: 1, kind: 'reply', id: 1, ok: true,
      value: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, sheets: [], activeSheet: 0 },
    });
    await expect(opening).rejects.toMatchObject({
      name: 'SessionFailure', code: 'disposed', message: 'Session was disposed',
    });
    expect(closed).toBeGreaterThan(0);
  });
});
