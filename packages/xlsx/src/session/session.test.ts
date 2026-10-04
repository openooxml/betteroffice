import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { runInNewContext } from 'node:vm';
import {
  createSessionClient, createSessionHost, SessionFailure, type MethodPolicy, type SessionTransport,
} from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { XlsxEditRequest, XlsxRangeTarget, XlsxReadResult } from '../edits';
import { initWasm, openWorkbook } from '../wasm/loader';
import { createWorkbookSession, type WorkbookSession } from './client';
import { createWorkbookSessionHost } from './host';
import {
  WORKBOOK_SESSION_METHODS,
  WORKBOOK_SESSION_POLICIES,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
} from './methods';

let fixture: Uint8Array;
let wasmBytes: Uint8Array<ArrayBuffer>;

beforeAll(async () => {
  const [wasm, xlsx] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/xlsx_wasm_bg.wasm')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')),
  ]);
  wasmBytes = new Uint8Array(wasm);
  await initWasm(wasmBytes);
  fixture = new Uint8Array(xlsx);
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

async function session(): Promise<WorkbookSession> {
  const pair = createInProcessPair();
  createWorkbookSessionHost(pair.host);
  return createWorkbookSession(fixture, {}, pair.client);
}

describe('workbook sessions', () => {
  test('declares ordered methods and marks edits as user input', () => {
    expect(Object.keys(WORKBOOK_SESSION_METHODS)).toEqual(Object.keys(WORKBOOK_SESSION_POLICIES));
    expect(WORKBOOK_SESSION_POLICIES.applyEdits).toEqual({
      lane: 'input', mutates: true, userInput: true,
    });
    for (const policy of Object.values(WORKBOOK_SESSION_POLICIES)) {
      expect(typeof policy).toBe('object');
      expect((policy as MethodPolicy).reorderable).not.toBe(true);
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
      const request: XlsxEditRequest = {
        expectVersion: workerBefore.version,
        steps: [
          { op: 'setCellInputs', target: target('B3'), inputs: [['1000']] },
          { op: 'patchStyle', target: target('B3'), patch: { bold: true } },
          { op: 'setNumberFormat', target: target('C3'), format: 'percent' },
        ],
      };
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
});
