import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { isTransferResult } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { initWasm } from '../wasm/loader';
import { DEFAULT_MAX_FRAME_BYTES } from '../collaboration/protocol';
import type { WorkbookHandle } from '../wasm/loader';
import { createWorkbookCommandState } from './commandHandlers';
import { createWorkbookCommandHost } from './commandHost';
import type { WorkbookSharedCommand } from './commandProtocol';
import { createWorkbookSessionHost } from './host';
import { WORKBOOK_SESSION_METHODS, type WorkbookSessionEvents, type WorkbookSessionMethods } from './methods';
import { batchRequests } from './testHelpers';

function pipeFrameHost(changes: boolean[] = []) {
  const identity = { session: 'book', incarnation: 'worker' };
  let version = 0;
  let frames = 0;
  const handle = {
    version: () => `v${version}`, sheetCount: () => 2,
    sheetInfo: () => ({ sheetIds: ['sheet:0', 'sheet:1'], sheetNames: ['First', 'Second'], activeSheet: 0 }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }, { editable: true }] }),
    historyState: () => ({ canUndo: false, canRedo: false }), listProposals: () => [],
    validateEdits: () => ({ ok: true, version: `v${version}` }), editCell: () => ({ applied: false }),
    selectionFormatting: () => ({}), mergedRanges: () => [], calculationStatus: () => ({}),
  } as unknown as WorkbookHandle;
  const state = createWorkbookCommandState(identity, { epoch: 1, version: handle.version() },
    { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0', 'sheet:1']);
  const host = createWorkbookCommandHost({ handle, state,
    applyPipe: () => { if (changes.shift()) version++; },
    frame: (request, coversThroughSeq) => {
      frames++;
      return { ...identity, coversThroughSeq, requestGeneration: request.requestGeneration ?? request.generation,
        zoom: request.zoom, dpr: request.dpr, version: handle.version(), epoch: frames,
        sequence: coversThroughSeq, sheet: request.sheet, viewport: request.viewport,
        displayList: new TextEncoder().encode('{"commands":[]}').buffer };
    } });
  const request = { sheet: 0, viewport: { x: 0, y: 0, width: 160, height: 64 }, generation: 1, zoom: 1, dpr: 1 };
  const command = (seq: number, frame?: WorkbookSharedCommand['payload']['frame']): WorkbookSharedCommand => ({
    session: 'book', incarnation: 1, seq, basis: { epoch: 1, version: 0 }, kind: 'collabPipe',
    payload: { payload: { provider: 1, ops: [{ op: 'connect', seq }] }, frame },
  });
  return { host, request, command, frames: () => frames };
}

test('xlsx pipe no-ops omit frames for an already represented version and viewport', async () => {
  const { host, request, command, frames } = pipeFrameHost();
  host.rememberFrame(request);
  for (const input of [command(1, { ...request, requestGeneration: 2 }), command(2)]) {
    const result = await host.handler(null, input);
    const ack = isTransferResult(result) ? result.value : result;
    expect(ack.outcomes).toMatchObject([{ status: 'noop', workbook: { committed: false } }]);
    expect(ack.frame).toBeUndefined();
    expect(ack.engineVersion).toBe('v0');
  }
  expect(frames()).toBe(0);
  const result = await host.frame();
  expect(isTransferResult(result) ? result.value.coversThroughSeq : result.coversThroughSeq).toBe(2);
  expect(frames()).toBe(1);
});

test('xlsx pipe no-ops render unrepresented viewports and changed versions', async () => {
  const { host, request, command, frames } = pipeFrameHost([false, false, false, false, true, false]);
  await host.requestFrame(request);
  const requests = [
    { ...request, viewport: { ...request.viewport, x: 80 } },
    { ...request, zoom: 2 }, { ...request, dpr: 2 }, { ...request, sheet: 1 }, request,
  ];
  for (const [index, frame] of requests.entries()) {
    const result = await host.handler(null, command(index + 1, frame));
    const ack = isTransferResult(result) ? result.value : result;
    expect(ack.frame?.payload).toMatchObject({ sheet: frame.sheet, viewport: frame.viewport,
      zoom: frame.zoom, dpr: frame.dpr, version: index === 4 ? 'v1' : 'v0' });
  }
  const result = await host.handler(null, command(6));
  expect(isTransferResult(result) ? result.value.frame : result.frame).toBeUndefined();
  expect(frames()).toBe(6);
});

test('xlsx trailing pipe no-op preserves an unframed batched change', async () => {
  const { host, request, command, frames } = pipeFrameHost([true, false]);
  await host.requestFrame(request);
  const first = command(1, request);
  const second = command(2);
  host.policy(first);
  host.policy(second);
  const results = await Promise.all([host.handler(null, first), host.handler(null, second)]);
  const [changed, unchanged] = results.map((result) => isTransferResult(result) ? result.value : result);
  expect(changed.frame).toBeUndefined();
  expect(unchanged.outcomes).toMatchObject([{ seq: 2, status: 'noop', workbook: { committed: false } }]);
  expect(unchanged.frame?.payload).toMatchObject({ coversThroughSeq: 2, version: 'v1' });
  expect(frames()).toBe(2);
});

test('xlsx trailing pipe no-op preserves an unframed local input acknowledgement', async () => {
  const { host, request, command, frames } = pipeFrameHost();
  await host.requestFrame(request);
  const first: WorkbookSharedCommand = { session: 'book', incarnation: 1, seq: 1,
    basis: { epoch: 1, version: 0 }, kind: 'editCell', payload: { frame: request,
      payload: { args: [0, 0, 0, 'same'], target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } } } } };
  const second = command(2);
  host.policy(first);
  host.policy(second);
  const results = await Promise.all([host.handler(null, first), host.handler(null, second)]);
  const [local, pipe] = results.map((result) => isTransferResult(result) ? result.value : result);
  expect(local.outcomes).toMatchObject([{ status: 'noop' }]);
  expect(local.frame).toBeUndefined();
  expect(pipe.frame?.payload).toMatchObject({ coversThroughSeq: 2, version: 'v0' });
  expect(frames()).toBe(2);
});

test('xlsx shared host retains each provider frame limit until destroy', async () => {
  const identity = { session: 'book', incarnation: 'opaque:worker' };
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const handle = {
    version: () => 'opaque:initial', sheetCount: () => 1,
    sheetInfo: () => ({ sheetIds: ['sheet:0'], sheetNames: ['Data'], activeSheet: 0 }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }] }),
    historyState: () => ({ canUndo: false, canRedo: false }),
    selectionFormatting: () => ({}), mergedRanges: () => [], calculationStatus: () => ({}),
  } as unknown as WorkbookHandle;
  const state = createWorkbookCommandState(identity, { epoch: 1, version: handle.version() }, selected, ['sheet:0']);
  const applied: unknown[] = [];
  const host = createWorkbookCommandHost({ handle, state, applyPipe: (payload) => { applied.push(payload); },
    frame: () => { throw new Error('Unexpected frame'); } });
  const maxFrameBytes = DEFAULT_MAX_FRAME_BYTES + 1024;
  const payloads = [
    { provider: 1, ops: [{ op: 'create', seq: 1, options: { maxFrameBytes } }] },
    { provider: 2, ops: [{ op: 'create', seq: 1, options: { maxFrameBytes: 1024 } }] },
    { provider: 1, ops: [{ op: 'event', seq: 2, event: { type: 'message', data: new ArrayBuffer(maxFrameBytes) } }] },
    { provider: 2, ops: [{ op: 'event', seq: 2, event: { type: 'message', data: new ArrayBuffer(1025) } }] },
    { provider: 1, ops: [{ op: 'destroy', seq: 3 }] },
    { provider: 1, ops: [{ op: 'create', seq: 4, options: {} },
      { op: 'event', seq: 5, event: { type: 'message', data: new ArrayBuffer(maxFrameBytes) } }] },
  ];
  for (const [index, payload] of payloads.entries()) {
    const result = await host.handler(null, { session: 'book', incarnation: 1, seq: index + 1,
      basis: { epoch: 1, version: 0 }, kind: 'collabPipe', payload: { payload } } as WorkbookSharedCommand);
    const ack = isTransferResult(result) ? result.value : result;
    expect(ack.outcomes).toMatchObject([{ seq: index + 1, status: index === 3 || index === 5 ? 'refused' : 'noop',
      ...(index === 3 || index === 5 ? { reason: 'invalid-command' } : {}) }]);
  }
  expect(applied).toHaveLength(4);
  expect(state.pipeFrameBytes?.has(1)).toBe(false);
});

test('xlsx shared host omits invalid viewport frames and remains usable for ordered frames', async () => {
  const identity = { session: 'book', incarnation: 'opaque:worker' };
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const handle = {
    version: () => 'opaque:initial', sheetCount: () => 1,
    sheetInfo: () => ({ sheetIds: ['sheet:0'], sheetNames: ['Data'], activeSheet: 0 }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }] }),
    historyState: () => ({ canUndo: false, canRedo: false }),
    selectionFormatting: () => ({}), mergedRanges: () => [], calculationStatus: () => ({}),
  } as unknown as WorkbookHandle;
  const state = createWorkbookCommandState(identity, { epoch: 1, version: handle.version() }, selected, ['sheet:0']);
  let pixels = 0;
  const host = createWorkbookCommandHost({ handle, state, frame: (request, coversThroughSeq) => {
    pixels += 1;
    return { ...identity, coversThroughSeq, requestGeneration: request.generation,
      zoom: request.zoom, dpr: request.dpr, version: handle.version(), epoch: pixels,
      sequence: coversThroughSeq, sheet: request.sheet, viewport: request.viewport,
      displayList: new TextEncoder().encode('{"commands":[]}').buffer };
  } });
  function command(seq: number, frame?: WorkbookSharedCommand['payload']['frame']): WorkbookSharedCommand {
    return { session: 'book', incarnation: 1, basis: { epoch: 1, version: 0 }, seq,
      kind: 'setReadOnly', payload: { payload: { readOnly: seq === 1 }, frame } };
  }
  const first = await host.handler(null, command(1));
  const ack = isTransferResult(first) ? first.value : first;
  expect(ack).toMatchObject({ throughSeq: 1, version: 0, engineVersion: 'opaque:initial' });
  expect(ack.frame).toBeUndefined();
  expect(pixels).toBe(0);
  await expect(host.frame()).rejects.toThrow('No frame');
  const invalid = command(2, { sheet: 0, viewport: { x: 0, y: 0, width: 0, height: 40 },
    generation: 1, zoom: 1, dpr: 1 });
  const second = await host.handler(null, invalid);
  expect(isTransferResult(second) ? second.value.frame : second.frame).toBeUndefined();
  expect(pixels).toBe(0);
  const third = await host.handler(null, command(3));
  expect(isTransferResult(third) ? third.value.throughSeq : third.throughSeq).toBe(3);
  const request = { sheet: 0, viewport: { x: 0, y: 0, width: 80, height: 40 }, generation: 2, zoom: 1, dpr: 1 };
  const rendered = await host.requestFrame(request);
  const frame = isTransferResult(rendered) ? rendered.value : rendered;
  expect(frame).toMatchObject({ coversThroughSeq: 3, revision: 1, payload: { viewport: request.viewport } });
  expect(pixels).toBe(1);
  const later = await host.handler(null, command(4));
  expect(isTransferResult(later) ? later.value.throughSeq : later.throughSeq).toBe(4);
});

test('ordered xlsx command frames release the resident lane for later admitted edits', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const pair = createInProcessPair();
  createWorkbookSessionHost(batchRequests(pair.host));
  const client = createSessionClient<WorkbookSessionMethods,
    { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] }>(pair.client, {
    methods: WORKBOOK_SESSION_METHODS,
  });
  try {
    const opened = await client.call.open(bytes.slice().buffer as ArrayBuffer);
    const identity = opened.commandContext!;
    const target = { sheetId: opened.sheets[0].id, range: { kind: 'a1' as const, a1: 'A1' } };
    const request = { sheet: 0, viewport: { x: 0, y: 0, width: 160, height: 64 },
      generation: 9, zoom: 1, dpr: 1 };
    const command = (seq: number): WorkbookSharedCommand => ({
      session: identity.session, incarnation: 1, seq, basis: { epoch: identity.basis.epoch, version: 0 },
      ...(seq === 2 ? { afterSeq: 1 } : {}), kind: 'editCell',
      payload: { payload: { args: [0, 0, 0, String(seq)], target } },
    });
    const first = client.call.edit(command(1));
    const framing = client.call.commandFrame(request);
    const second = client.call.edit(command(2));
    const [before, frame, after] = await Promise.all([first, framing, second]);
    expect(before.outcomes).toMatchObject([{ seq: 1, status: 'applied' }]);
    expect(after.outcomes).toMatchObject([{ seq: 2, status: 'applied' }]);
    expect(frame).toMatchObject({ coversThroughSeq: 2, payload: { sequence: 2,
      version: after.engineVersion, requestGeneration: 9, viewport: request.viewport } });
    const displayList = JSON.parse(new TextDecoder().decode(frame!.payload.displayList));
    expect(displayList.commands.some((entry: { op: string; text?: string }) => entry.op === 'text' && entry.text === '2')).toBe(true);
    expect((await client.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('2');
  } finally {
    try { await client.call.dispose(); } finally { await client.dispose(); }
  }
});

test('large unchanged selection projections copy no edit state and changes stay sparse', async () => {
  const identity = { session: 'book', incarnation: 'opaque:worker' };
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const mergedRanges = Array.from({ length: 5000 }, (_, row) => ({ start: { row, col: 0 }, end: { row, col: 1 } }));
  const handle = {
    version: () => 'opaque:initial', sheetCount: () => 1,
    sheetInfo: () => ({ sheetIds: ['sheet:0'], sheetNames: ['Data'], activeSheet: 0 }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }] }),
    historyState: () => ({ canUndo: false, canRedo: false }),
    selectionFormatting: () => ({}), mergedRanges: () => mergedRanges,
    calculationStatus: () => ({ limitedCells: [] }),
  } as unknown as WorkbookHandle;
  const state = createWorkbookCommandState(identity, { epoch: 1, version: handle.version() }, selected, ['sheet:0']);
  const host = createWorkbookCommandHost({ handle, state,
    frame: () => { throw new Error('Unexpected frame'); } });
  const select = (seq: number): WorkbookSharedCommand => ({ session: 'book', incarnation: 1,
    basis: { epoch: 1, version: 0 }, seq, kind: 'select', payload: { payload: { selection: selected } } });
  const first = await host.handler(null, select(1));
  const initial = isTransferResult(first) ? first.value : first;
  expect(initial.editState?.mergedRanges).toEqual(mergedRanges);
  expect(initial.compatibility?.editState?.mergedRanges).toEqual(mergedRanges);
  const clone = globalThis.structuredClone;
  const copied: string[] = [];
  const spy = spyOn(globalThis, 'structuredClone').mockImplementation((value, options) => {
    copied.push(JSON.stringify(value));
    return clone(value, options);
  });
  try {
    const repeated = await host.handler(null, select(2));
    const unchanged = isTransferResult(repeated) ? repeated.value : repeated;
    expect(unchanged).toMatchObject({ throughSeq: 2, outcomes: [{ seq: 2, status: 'noop' }] });
    expect(unchanged.selection).toBeUndefined();
    expect(unchanged.editState).toBeUndefined();
    expect(unchanged.compatibility?.selection).toBeUndefined();
    expect(unchanged.compatibility?.editState).toBeUndefined();
    expect(copied).toEqual([JSON.stringify(selected), JSON.stringify({ seq: 2, status: 'noop', committed: false })]);
    expect(copied.reduce((bytes, value) => bytes + new TextEncoder().encode(value).byteLength, 0))
      .toBe(new TextEncoder().encode(JSON.stringify(selected)).byteLength +
        new TextEncoder().encode(JSON.stringify({ seq: 2, status: 'noop', committed: false })).byteLength);
    copied.length = 0;
    const result = await host.handler(null, { session: 'book', incarnation: 1, seq: 3,
      basis: { epoch: 1, version: 0 }, kind: 'setReadOnly', payload: { payload: { readOnly: true } } });
    const changed = isTransferResult(result) ? result.value : result;
    expect(changed.outcomes).toMatchObject([{ seq: 3, status: 'applied', committed: false }]);
    expect(changed.version).toBe(0);
    expect(changed.editState).toEqual({ readOnly: true });
    expect(changed.compatibility?.editState).toEqual({ readOnly: true });
    expect(copied).toEqual([JSON.stringify({ seq: 3, status: 'applied', committed: false }),
      JSON.stringify({ readOnly: true }), JSON.stringify({ readOnly: true })]);
    expect(state.editState?.mergedRanges).toEqual(mergedRanges);
  } finally { spy.mockRestore(); }
});
