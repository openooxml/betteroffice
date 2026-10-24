import { expect, mock, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { isTransferResult } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { initWasm } from '../wasm/loader';
import type { WorkbookHandle } from '../wasm/loader';
import { createWorkbookCommandState, projectEditState } from './commandHandlers';
import { createWorkbookCommandHost } from './commandHost';
import type { CommandStatus, WorkbookCommand, WorkbookSharedCommand, WorkbookSharedWireAck } from './commandProtocol';
import {
  decodeWorkbookRecovery, encodeWorkbookRecovery, isCollaborativeWorkbookRecovery, restoreWorkbookCommandState,
} from './recovery';
import { createWorkbookSessionHost } from './host';
import { WORKBOOK_SESSION_METHODS, type WorkbookSessionEvents, type WorkbookSessionMethods } from './methods';
import { batchRequests } from './testHelpers';

async function workbookBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  return zip.generateAsync({ type: 'arraybuffer' });
}

function session(batched = false) {
  const pair = createInProcessPair();
  createWorkbookSessionHost(batched ? batchRequests(pair.host) : pair.host);
  return createSessionClient<WorkbookSessionMethods,
    { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] }>(pair.client, { methods: WORKBOOK_SESSION_METHODS });
}

test('checkpoint after N commands restores a fresh session boundary', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true,
      calculation: { nowSerial: 45000, randSeed: 19 } });
    const context = opened.commandContext!;
    const target = { sheetId: opened.sheets[0].id, range: { kind: 'a1' as const, a1: 'A1' } };
    const input = (seq: number, kind: string, payload: unknown): WorkbookSharedCommand => ({
      session: context.session, incarnation: 1, seq, basis: { epoch: context.basis.epoch, version: 0 },
      kind, payload: { payload: payload as WorkbookSharedCommand['payload']['payload'] },
    });
    const edit = await original.call.edit(input(1, 'editCell', { args: [0, 0, 0, '=7*2'], target }));
    const format = await original.call.edit(input(2, 'applyOps', { toolbar: { id: 'bold', args: null, target } }));
    const selection = { ...target, range: { kind: 'a1' as const, a1: 'B2' } };
    const selected = await original.call.edit(input(3, 'select', { selection }));
    const mode = await original.call.edit(input(4, 'setReadOnly', { readOnly: true }));
    const projection = { ...opened.editState!, ...edit.editState, ...format.editState, ...selected.editState, ...mode.editState };
    const saved = await original.call.save();
    const checkpoint = await original.call.checkpoint();
    expect(decodeWorkbookRecovery(checkpoint.state).engine[8]).toBe(7);
    expect(await original.call.version()).toBe(mode.engineVersion);
    await expect(replacement.call.openRecovery({ ...checkpoint, throughSeq: checkpoint.throughSeq - 1 }))
      .rejects.toThrow('envelope mismatch');
    const corrupt = checkpoint.state.slice();
    corrupt[corrupt.length - 1] = (corrupt[corrupt.length - 1] ?? 0) ^ 1;
    await expect(replacement.call.openRecovery({ ...checkpoint, state: corrupt })).rejects.toThrow();
    const restored = await replacement.call.openRecovery(checkpoint);
    expect(restored).toMatchObject({ throughSeq: 4, revision: checkpoint.version, historyReset: true,
      version: 2, dirty: true, selection, sheets: projection.sheets,
      editState: { ...projection, canUndo: false, canRedo: false, historyReset: true } });
    expect(restored.commandContext).toMatchObject({ session: context.session, basis: { epoch: context.basis.epoch,
      version: mode.engineVersion } });
    expect(restored.commandContext!.incarnation).not.toBe(context.incarnation);
    expect(await replacement.call.version()).toBe(mode.engineVersion);
    expect(await replacement.call.save()).toEqual(saved);
    for (const seq of [1, 4]) {
      const duplicate = await replacement.call.edit({ ...input(seq, 'editCell', { args: [0, 0, 0, '99'], target }), replay: true });
      expect(duplicate).toMatchObject({ throughSeq: 4, version: checkpoint.version, outcomes: [{ seq, status: 'duplicate' }] });
    }
    const next = await replacement.call.edit({ ...input(5, 'setReadOnly', { readOnly: false }),
      basis: { epoch: context.basis.epoch, version: checkpoint.version } });
    expect(next).toMatchObject({ throughSeq: 5, outcomes: [{ seq: 5, status: 'applied', committed: false }],
      editState: { canUndo: false, canRedo: false, historyReset: true, readOnly: false } });
    expect(await replacement.call.save()).toEqual(saved);
    const continued = await replacement.call.edit({ ...input(6, 'editCell', { args: [0, 1, 1, '21'], target: selection }),
      basis: { epoch: context.basis.epoch, version: checkpoint.version }, afterSeq: 5 });
    expect(continued).toMatchObject({ throughSeq: 6, version: checkpoint.version + 1,
      outcomes: [{ seq: 6, status: 'applied' }], editState: { canUndo: true } });
    expect((await replacement.call.cellInputs(0, 'B2')).cells[0][0].input).toBe('21');
  } finally { await original.dispose(); await replacement.dispose(); }
});

function handle(version = 'engine:0') {
  return {
    version: () => version, sheetCount: () => 1,
    sheetInfo: () => ({ sheetIds: ['sheet:0'], sheetNames: ['Data'], activeSheet: 0 }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }] }),
    historyState: () => ({ canUndo: false, canRedo: false }),
    selectionFormatting: () => ({}), mergedRanges: () => [], calculationStatus: () => ({}),
  } as unknown as WorkbookHandle;
}

function engineHeader(metadata = '{"collaborative":false}', version = 3) {
  const payload = new TextEncoder().encode(metadata);
  const engine = new Uint8Array(42 + payload.length);
  engine.set([88, 76, 83, 88, 82, 69, 67, 86, version]);
  engine[41] = payload.length;
  engine.set(payload, 42);
  return engine;
}

test('recovery reads engine checkpoints from before and since patterned fills', () => {
  for (const version of [3, 5]) expect(isCollaborativeWorkbookRecovery(engineHeader(undefined, version))).toBe(false);
});

test.each([
  ['unsupported codec', engineHeader(undefined, 4)],
  ['invalid header', new Uint8Array(42)],
  ['truncated header', engineHeader().slice(0, 41)],
  ['truncated metadata', engineHeader().slice(0, -1)],
  ['unterminated metadata length', new Uint8Array([...engineHeader().slice(0, 41), 128, 128, 128, 128, 128])],
  ['malformed metadata', engineHeader('{')],
  ['null metadata', engineHeader('null')],
  ['missing collaboration mode', engineHeader('{}')],
  ['invalid collaboration mode', engineHeader('{"collaborative":"false"}')],
  ['invalid metadata encoding', engineHeader('{"collaborative":false,"value":"x"}').map((byte) => byte === 120 ? 255 : byte)],
] as const)('recovery refuses %s before initialization', async (_name, engine) => {
  const opened = handle();
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0']);
  const checkpoint = { epoch: 1, version: 0, throughSeq: 0,
    state: encodeWorkbookRecovery(engine, state, 0, [[0, opened.version()]]) };
  const pair = createInProcessPair();
  const initialize = mock(async () => { throw new Error('Unexpected recovery initialization'); });
  createWorkbookSessionHost(pair.host, { initWasm: initialize });
  const replacement = createSessionClient<WorkbookSessionMethods, {}>(pair.client, { methods: WORKBOOK_SESSION_METHODS });
  try {
    await expect(replacement.call.openRecovery(checkpoint)).rejects.toThrow('Unsupported workbook recovery checkpoint');
    expect(initialize).not.toHaveBeenCalled();
    expect(replacement.failure).toBeUndefined();
  } finally { await replacement.dispose(); }
});

test('checkpoint measures each synchronous encode with an injected clock', async () => {
  const opened = handle();
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0']);
  let now = 0;
  let duration = 20;
  const options = { handle: opened, state, now: () => now,
    captureRecovery: () => { now += duration; return new Uint8Array([1]); },
    frame: () => { throw new Error('Unexpected frame'); } };
  const host = createWorkbookCommandHost(options);
  for (duration of [20, 120, 40]) {
    const result = await host.checkpoint(null, true);
    const checkpoint = isTransferResult(result) ? result.value : result;
    expect(checkpoint).toMatchObject({ encodeMs: duration, throughSeq: 0 });
  }
});

test('checkpoint restores edited state and command metadata without a peer', async () => {
  const opened = handle('engine:edited');
  const identity = { session: 'book', incarnation: 'first' };
  const selected = { sheetId: 'command-sheet:4', range: { kind: 'a1' as const, a1: 'B2' } };
  const state = createWorkbookCommandState(identity, { epoch: 3, version: opened.version() }, selected, [selected.sheetId]);
  state.nextSheetId = 5;
  state.calculation = { nowSerial: 45000, randSeed: 12 };
  state.formats.set('format:2', { rows: 1, columns: 1, formats: [{}] });
  const engine = new Uint8Array([7, 8, 9]);
  const host = createWorkbookCommandHost({ handle: opened, state, revision: 9,
    versions: [[9, opened.version()]], captureRecovery: () => engine,
    frame: () => { throw new Error('Unexpected frame'); } });
  const command = (seq: number): WorkbookSharedCommand => ({ session: 'book', incarnation: 1, seq,
    basis: { epoch: 3, version: 9 }, kind: 'select', payload: { payload: { selection: selected } } });
  await host.handler(null, command(1));
  await host.handler(null, command(2));
  const captured = await host.checkpoint(null);
  const checkpoint = isTransferResult(captured) ? captured.value : captured;
  expect(checkpoint).toMatchObject({ throughSeq: 2, version: 9, epoch: 3 });
  const recovery = decodeWorkbookRecovery(checkpoint.state);
  expect(recovery.engine).toEqual(engine);
  const restored = restoreWorkbookCommandState(opened, recovery, { session: 'book', incarnation: 'replacement' });
  expect(restored).toMatchObject({ throughSeq: 2, nextSheetId: 5, calculation: state.calculation,
    selection: selected, readOnly: false, historyReset: true, historySelections: [] });
  expect(restored.formats).toEqual(state.formats);
  if (!state.editState) throw new Error('Missing captured edit state');
  expect(restored.editState).toEqual({ ...state.editState, historyReset: true });
  const fresh = createWorkbookCommandHost({ handle: opened, state: restored, revision: recovery.revision,
    versions: recovery.versions, captureRecovery: () => engine,
    frame: () => { throw new Error('Unexpected frame'); } });
  for (const seq of [1, 2]) {
    const replay = await fresh.handler(null, { ...command(seq), replay: true });
    const ack = isTransferResult(replay) ? replay.value : replay;
    expect(ack.throughSeq).toBe(2);
    expect(ack.outcomes[0].status).toBe('duplicate');
  }
  const next = await fresh.handler(null, command(3));
  const ack = isTransferResult(next) ? next.value : next;
  expect(ack).toMatchObject({ throughSeq: 3, version: 9, engineVersion: 'engine:edited',
    editState: { canUndo: false, canRedo: false, historyReset: true } });
});

test('recovery refuses corrupt metadata before command admission', () => {
  const opened = handle();
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0']);
  projectEditState(opened, state);
  const bytes = encodeWorkbookRecovery(new Uint8Array([1]), state, 0, [[0, opened.version()]]);
  expect(() => decodeWorkbookRecovery(bytes.slice(0, -1))).toThrow();
  const malformed = bytes.slice();
  malformed[8] = 99;
  expect(() => decodeWorkbookRecovery(malformed)).toThrow();
  const corrupt = bytes.slice();
  corrupt[corrupt.length - 1] = (corrupt[corrupt.length - 1] ?? 0) ^ 1;
  expect(() => decodeWorkbookRecovery(corrupt)).toThrow();
  expect(state.throughSeq).toBe(0);
  expect(state.statuses.size).toBe(0);
});

test.each<[string, number[], CommandStatus | undefined]>([
  ['zero', [0], undefined],
  ['non-integer', [1.5], undefined],
  ['unsafe integer', [Number.MAX_SAFE_INTEGER + 1], undefined],
  ['missing retained status', [2], undefined],
  ['applied retained status', [2], { seq: 2, status: 'applied', committed: true }],
  ['noop retained status', [2], { seq: 2, status: 'noop', committed: false }],
  ['cancelled retained status', [2], { seq: 2, status: 'cancelled', reason: 'target-sheet-missing', committed: false }],
  ['different retained reason', [2], { seq: 2, status: 'refused', reason: 'stale-version', committed: false }],
  ['committed retained refusal', [2], { seq: 2, status: 'refused', reason: 'target-sheet-missing', committed: true }],
])('recovery rejects refused outcome affected sequences with %s', (_, affectedSeqs, retained) => {
  const opened = handle();
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0']);
  state.throughSeq = 2;
  const outcome = { seq: 1, status: 'refused' as const, reason: 'target-sheet-missing', committed: false as const,
    affectedSeqs };
  state.statuses.set(1, { ...outcome, affectedSeqs: [1] });
  if (retained) state.statuses.set(2, retained);
  const refusedSeqs = [...state.statuses].filter(([, status]) => status.status === 'refused' || status.status === 'cancelled')
    .map(([seq]) => seq);
  const bytes = encodeWorkbookRecovery(new Uint8Array([1]), state, 0, [[0, opened.version()]],
    refusedSeqs, undefined, undefined, { refusedOutcomes: [[1, outcome]] });
  expect(() => decodeWorkbookRecovery(bytes)).toThrow(new TypeError('Invalid workbook recovery checkpoint'));
});

test('recovery stores engine bytes without JSON expansion', () => {
  const opened = handle();
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0']);
  const engine = new Uint8Array(1024 * 1024).fill(255);
  const bytes = encodeWorkbookRecovery(engine, state, 0, [[0, opened.version()]]);
  expect(bytes.length).toBeLessThan(engine.length + 1024);
  const decoded = decodeWorkbookRecovery(bytes);
  expect(decoded.engine).toEqual(engine);
  expect(decoded.engine.buffer).toBe(bytes.buffer);
});

test('engine checkpoints are read at codecs 3, 5, 6 and 7 and refused at any other', () => {
  const engine = (codec: number, collaborative: boolean) => {
    const metadata = new TextEncoder().encode(JSON.stringify({ collaborative }));
    return new Uint8Array([...new TextEncoder().encode('XLSXRECV'), codec, ...new Uint8Array(32), metadata.length,
      ...metadata]);
  };
  for (const codec of [3, 5, 6, 7]) {
    expect(isCollaborativeWorkbookRecovery(engine(codec, false))).toBe(false);
    expect(isCollaborativeWorkbookRecovery(engine(codec, true))).toBe(true);
  }
  for (const codec of [2, 4, 8]) {
    expect(() => isCollaborativeWorkbookRecovery(engine(codec, false))).toThrow('Unsupported workbook recovery checkpoint');
  }
});

test('recovery preserves applied dependencies and compatibility duplicates', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true });
    const context = opened.commandContext!;
    const target = { sheetId: opened.sheets[0]!.id, range: { kind: 'a1' as const, a1: 'A1' } };
    const command: WorkbookCommand = { session: context.session, incarnation: context.incarnation, seq: 1,
      basis: context.basis, kind: 'editCell', payload: { args: [0, 0, 0, '7'], target } };
    const applied = await original.call.command(command);
    expect(applied.statuses[0]).toMatchObject({ seq: 1, status: 'applied', committed: true });
    const checkpoint = await original.call.checkpoint();
    const restored = await replacement.call.openRecovery(checkpoint);
    const restoredCommand = { ...command, incarnation: restored.commandContext!.incarnation };
    const duplicate = await replacement.call.command({ ...restoredCommand, payload: { ...command.payload,
      args: [0, 0, 0, '99'] } });
    expect(duplicate.statuses).toEqual(applied.statuses);
    expect(duplicate.throughSeq).toBe(1);
    const dependent: WorkbookCommand = { ...restoredCommand, seq: 2, afterSeq: 1,
      payload: { ...command.payload, args: [0, 0, 0, '8'] } };
    const uninterrupted = await original.call.command({ ...dependent, incarnation: context.incarnation });
    const recovered = await replacement.call.command(dependent);
    expect(recovered.statuses).toEqual(uninterrupted.statuses);
    expect(recovered.statuses[0]).toMatchObject({ seq: 2, status: 'applied', committed: true });
    expect(await replacement.call.save()).toEqual(await original.call.save());
  } finally { await original.dispose(); await replacement.dispose(); }
});

test('recovery excludes identity refusals and stale epoch duplicates from its ledger', async () => {
  const opened = handle();
  const identity = { session: 'book', incarnation: 'first' };
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const state = createWorkbookCommandState(identity, { epoch: 1, version: opened.version() }, selected, ['sheet:0']);
  const options = { handle: opened, state, captureRecovery: () => new Uint8Array([1]),
    frame: () => { throw new Error('Unexpected frame'); } };
  const host = createWorkbookCommandHost(options);
  const command: WorkbookSharedCommand = { session: 'book', incarnation: 1, seq: 1,
    basis: { epoch: 1, version: 0 }, kind: 'select', payload: { payload: { selection: selected } } };
  for (const invalid of [{ ...command, session: 'other' }, { ...command, incarnation: 2 }]) {
    const result = await host.handler(null, invalid);
    expect(isTransferResult(result) ? result.value : result).toMatchObject({ throughSeq: 0,
      outcomes: [{ seq: 1, status: 'refused' }] });
    const captured = await host.checkpoint(null);
    const checkpoint = isTransferResult(captured) ? captured.value : captured;
    const recovery = decodeWorkbookRecovery(checkpoint.state);
    expect(recovery.refusedSeqs).toEqual([]);
    const restored = restoreWorkbookCommandState(opened, recovery, { ...identity, incarnation: 'replacement' });
    const fresh = createWorkbookCommandHost({ ...options, state: restored,
      revision: recovery.revision, versions: recovery.versions, refusedSeqs: recovery.refusedSeqs });
    const valid = await fresh.handler(null, command);
    expect(isTransferResult(valid) ? valid.value : valid).toMatchObject({ throughSeq: 1,
      outcomes: [{ seq: 1, status: 'noop' }] });
    await fresh.handler(null, { ...command, basis: { epoch: 0, version: 0 } });
    const recaptured = await fresh.checkpoint(null);
    const next = isTransferResult(recaptured) ? recaptured.value : recaptured;
    const recovered = decodeWorkbookRecovery(next.state);
    expect(recovered.refusedSeqs).toEqual([]);
    expect(recovered.statuses).toEqual([[1, { seq: 1, status: 'noop', committed: false }]]);
    expect(restoreWorkbookCommandState(opened, recovered, identity).throughSeq).toBe(1);
  }
});

test('wrong-session checkpoint restores and accepts the next valid edit', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true });
    const context = opened.commandContext!;
    const command: WorkbookSharedCommand = { session: context.session, incarnation: 1, seq: 1,
      basis: { epoch: context.basis.epoch, version: 0 }, kind: 'editCell',
      payload: { payload: { args: [0, 0, 0, '7'],
        target: { sheetId: opened.sheets[0]!.id, range: { kind: 'a1', a1: 'A1' } } } } };
    const refused = await original.call.edit({ ...command, session: 'other' });
    expect(refused).toMatchObject({ throughSeq: 0, outcomes: [{ seq: 1, status: 'refused', reason: 'wrong-session' }] });
    const checkpoint = await original.call.checkpoint();
    expect(checkpoint.throughSeq).toBe(0);
    const restored = await replacement.call.openRecovery(checkpoint);
    expect(restored.throughSeq).toBe(0);
    const applied = await replacement.call.edit(command);
    expect(applied).toMatchObject({ throughSeq: 1, outcomes: [{ seq: 1, status: 'applied' }] });
    expect((await replacement.call.cellInputs(0, 'A1')).cells[0]?.[0]?.input).toBe('7');
  } finally { await original.dispose(); await replacement.dispose(); }
});

test('checkpoint fences refusals and restores dependency admission', async () => {
  const opened = handle();
  const identity = { session: 'book', incarnation: 'first' };
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const state = createWorkbookCommandState(identity, { epoch: 1, version: opened.version() }, selected, ['sheet:0']);
  const options = { handle: opened, state, captureRecovery: () => new Uint8Array([1]),
    frame: () => { throw new Error('Unexpected frame'); } };
  const host = createWorkbookCommandHost(options);
  const command = (seq: number): WorkbookSharedCommand => ({ session: 'book', incarnation: 1, seq,
    basis: { epoch: 1, version: 0 }, kind: 'select', payload: { payload: { selection: selected } } });
  const refused = host.handler(null, { ...command(1), payload: { payload: { selection: { ...selected, sheetId: 'missing' } } } });
  const capture = host.checkpoint(null);
  const later = host.handler(null, command(2));
  await refused;
  const result = await capture;
  const checkpoint = isTransferResult(result) ? result.value : result;
  expect(checkpoint.throughSeq).toBe(1);
  const recovery = decodeWorkbookRecovery(checkpoint.state);
  expect(recovery.throughSeq).toBe(1);
  expect(recovery.refusedSeqs).toEqual([1]);
  await later;
  const restored = restoreWorkbookCommandState(opened, recovery, { ...identity, incarnation: 'replacement' });
  const fresh = createWorkbookCommandHost({ ...options, state: restored,
    revision: recovery.revision, versions: recovery.versions, refusedSeqs: recovery.refusedSeqs });
  const dependent = await fresh.handler(null, { ...command(2), afterSeq: 1 });
  expect(isTransferResult(dependent) ? dependent.value : dependent).toMatchObject({ throughSeq: 2,
    outcomes: [{ seq: 2, status: 'cancelled', reason: 'dependency-refused' }] });
  const independent = await fresh.handler(null, command(3));
  expect(isTransferResult(independent) ? independent.value : independent).toMatchObject({ throughSeq: 3,
    outcomes: [{ seq: 3, status: 'noop' }] });
});

test('checkpoint after a structural edit restores layout and saved bytes', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true });
    const context = opened.commandContext!;
    const target = { sheetId: opened.sheets[0]!.id, range: { kind: 'a1' as const, a1: 'A1' } };
    const command = (seq: number, kind: string, payload: WorkbookCommand['payload']): WorkbookSharedCommand => ({
      session: context.session, incarnation: 1, seq, basis: { epoch: context.basis.epoch, version: 0 },
      kind, payload: { payload },
    });
    await original.call.edit(command(1, 'editCell', { args: [0, 0, 0, '7'], target }));
    const inserted = await original.call.edit(command(2, 'applyOps', {
      args: [[{ type: 'insertRows', sheet: 0, at: 0, count: 2 }]], target,
    }));
    expect(inserted.outcomes[0]?.workbook).toMatchObject({ status: 'applied', committed: true });
    const saved = await original.call.save();
    const checkpoint = await original.call.checkpoint();
    const restored = await replacement.call.openRecovery(checkpoint);
    expect(restored).toMatchObject({ throughSeq: 2, historyReset: true,
      editState: { canUndo: false, canRedo: false, historyReset: true } });
    expect((await replacement.call.cellInputs(0, 'A3')).cells[0]?.[0]?.input).toBe('7');
    expect(await replacement.call.sheetView(0)).toEqual(await original.call.sheetView(0));
    expect(await replacement.call.save()).toEqual(saved);
    const next = { ...command(3, 'editCell', { args: [0, 3, 0, '9'],
      target: { ...target, range: { kind: 'a1' as const, a1: 'A4' } } }),
      basis: { epoch: context.basis.epoch, version: checkpoint.version } };
    for (const active of [original, replacement]) {
      const edited = await active.call.edit(next);
      expect(edited.outcomes[0]?.workbook).toMatchObject({ status: 'applied', committed: true });
    }
    expect(await replacement.call.save()).toEqual(await original.call.save());
  } finally { await original.dispose(); await replacement.dispose(); }
});

test('direct applyEdits before any command checkpoints the committed version and basis', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true });
    const target = { sheetId: opened.sheets[0]!.id, range: { kind: 'a1' as const, a1: 'A1' } };
    const edited = await original.call.applyEdits({ expectVersion: opened.commandContext!.basis.version,
      steps: [{ op: 'setCellInputs', target, inputs: [['7']] }] });
    expect(edited).toMatchObject({ ok: true, applied: true });
    const saved = await original.call.save();
    const checkpoint = await original.call.checkpoint();
    const recovery = decodeWorkbookRecovery(checkpoint.state);
    expect(checkpoint).toMatchObject({ throughSeq: 0, version: 1 });
    expect(recovery.engineVersion).toBe(await original.call.version());
    expect(recovery.versions).toContainEqual([checkpoint.version, recovery.engineVersion]);
    expect(recovery.bases).toContainEqual([recovery.engineVersion, 0]);
    const restored = await replacement.call.openRecovery(checkpoint);
    expect(restored).toMatchObject({ throughSeq: 0, version: 1, dirty: true, historyReset: true,
      editState: { canUndo: false, canRedo: false, historyReset: true } });
    expect((await replacement.call.cellInputs(0, 'A1')).cells[0]?.[0]?.input).toBe('7');
    expect(await replacement.call.save()).toEqual(saved);
    const context = restored.commandContext!;
    const next: WorkbookSharedCommand = { session: context.session, incarnation: 1, seq: 1,
      basis: { epoch: context.basis.epoch, version: checkpoint.version }, kind: 'editCell',
      payload: { payload: { args: [0, 0, 0, '8'], target } } };
    for (const active of [original, replacement]) {
      const applied = await active.call.edit(next);
      expect(applied.outcomes[0]?.workbook).toMatchObject({ status: 'applied', committed: true });
    }
    expect(await replacement.call.save()).toEqual(await original.call.save());
  } finally { await original.dispose(); await replacement.dispose(); }
});

test('restored shared duplicates return captured proposal format token and refusal results', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true });
    const context = opened.commandContext!;
    const target = { sheetId: opened.sheets[0]!.id, range: { kind: 'a1' as const, a1: 'A1' } };
    const command = (seq: number, kind: string, payload: WorkbookCommand['payload']): WorkbookSharedCommand => ({
      session: context.session, incarnation: 1, seq, basis: { epoch: context.basis.epoch, version: 0 },
      kind, payload: { payload },
    });
    const commands = [
      command(1, 'propose', { args: ['agent', 'review', [{ sheet: 0, row: 0, col: 0, input: '7', target }]] }),
      command(2, 'applyOps', { toolbar: { id: 'paintFormat', args: null, target } }),
      command(3, 'applyEdits', { args: [{ expectVersion: context.basis.version,
        steps: [{ op: 'setCellInputs', target, inputs: [['9']], expect: { cells: [[{ displayText: 'different' }]] } }] }] }),
    ];
    const captured: WorkbookSharedWireAck[] = [];
    for (const input of commands) captured.push(await original.call.edit(input));
    expect(captured[0]?.outcomes[0]?.workbook).toMatchObject({ status: 'applied', result: { id: expect.any(String) } });
    expect(captured[1]?.outcomes[0]?.workbook).toMatchObject({ result: { formatToken: expect.any(String) } });
    expect(captured[2]?.outcomes[0]?.workbook).toMatchObject({ status: 'refused', reason: 'content-mismatch',
      result: { ok: false, failure: { code: 'content-mismatch' } } });
    const checkpoint = await original.call.checkpoint();
    await replacement.call.openRecovery(checkpoint);
    const saved = await replacement.call.save();
    for (const [index, input] of commands.entries()) {
      const duplicate = await replacement.call.edit({ ...input, replay: true });
      expect(duplicate.outcomes[0]?.workbook).toEqual(captured[index]?.outcomes[0]?.workbook);
      expect(duplicate.compatibility?.statuses).toEqual(captured[index]?.compatibility?.statuses);
      expect(duplicate.throughSeq).toBe(3);
    }
    expect(await replacement.call.save()).toEqual(saved);
    const recaptured = decodeWorkbookRecovery((await replacement.call.checkpoint()).state);
    const recovery = decodeWorkbookRecovery(checkpoint.state);
    expect(recaptured.statuses).toEqual(recovery.statuses);
    expect(recaptured.formats).toEqual(recovery.formats);
  } finally { await original.dispose(); await replacement.dispose(); }
});

test('many refusals retain a bounded checkpoint ledger and duplicate results', async () => {
  const opened = handle();
  const identity = { session: 'book', incarnation: 'first' };
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const state = createWorkbookCommandState(identity, { epoch: 1, version: opened.version() }, selected, ['sheet:0']);
  const options = { handle: opened, state, captureRecovery: () => new Uint8Array([1]),
    frame: () => { throw new Error('Unexpected frame'); } };
  const host = createWorkbookCommandHost(options);
  const command = (seq: number): WorkbookSharedCommand => ({ session: 'book', incarnation: 1, seq,
    basis: { epoch: 1, version: 0 }, kind: 'select',
    payload: { payload: { selection: { ...selected, sheetId: 'missing' } } } });
  for (let seq = 1; seq <= 1024; seq += 1) {
    const result = await host.handler(null, command(seq));
    expect(isTransferResult(result) ? result.value : result).toMatchObject({ throughSeq: seq,
      outcomes: [{ workbook: { status: 'refused', reason: 'target-sheet-missing' } }] });
  }
  const result = await host.checkpoint(null);
  const checkpoint = isTransferResult(result) ? result.value : result;
  const recovery = decodeWorkbookRecovery(checkpoint.state);
  expect(state.statuses.size).toBe(256);
  expect(recovery.statuses).toHaveLength(256);
  expect(recovery.refusedSeqs).toHaveLength(256);
  expect(recovery.refusedSeqs).toEqual(Array.from({ length: 256 }, (_, index) => index + 769));
  expect(recovery.refusedRanges).toEqual([[1, 1024]]);
  expect(recovery.refusedOutcomes).toHaveLength(256);
  const restored = restoreWorkbookCommandState(opened, recovery, { ...identity, incarnation: 'replacement' });
  const fresh = createWorkbookCommandHost({ ...options, state: restored,
    revision: recovery.revision, versions: recovery.versions, refusedSeqs: recovery.refusedSeqs,
    refusedRanges: recovery.refusedRanges, refusedOutcomes: recovery.refusedOutcomes });
  for (const seq of [769, 1024]) {
    const duplicate = await fresh.handler(null, command(seq));
    const ack = isTransferResult(duplicate) ? duplicate.value : duplicate;
    expect(ack.outcomes[0]?.workbook).toEqual(state.statuses.get(seq));
  }
  const dependent = await fresh.handler(null, { ...command(1025), afterSeq: 769 });
  expect(isTransferResult(dependent) ? dependent.value : dependent).toMatchObject({ throughSeq: 1025,
    outcomes: [{ status: 'cancelled', reason: 'dependency-refused' }] });
  const recaptured = await fresh.checkpoint(null);
  const next = isTransferResult(recaptured) ? recaptured.value : recaptured;
  expect(decodeWorkbookRecovery(next.state).refusedSeqs).toHaveLength(256);
  expect(decodeWorkbookRecovery(next.state).refusedRanges).toEqual([[1, 1025]]);
  expect(decodeWorkbookRecovery(next.state).refusedOutcomes).toHaveLength(256);
});

test('checkpoint preserves refused dependencies beyond the status window', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session();
  const replacement = session();
  try {
    const opened = await original.call.open(await workbookBytes(), { includeEditState: true });
    const context = opened.commandContext!;
    const target = { sheetId: opened.sheets[0]!.id, range: { kind: 'a1' as const, a1: 'A1' } };
    const command = (seq: number, kind: string, payload: WorkbookCommand['payload']): WorkbookSharedCommand => ({
      session: context.session, incarnation: 1, seq, basis: { epoch: context.basis.epoch, version: 0 },
      kind, payload: { payload },
    });
    const refused = await original.call.edit(command(1, 'select', { selection: { ...target, sheetId: 'missing' } }));
    expect(refused.outcomes).toMatchObject([{ seq: 1, status: 'refused', reason: 'target-sheet-missing' }]);
    for (let seq = 2; seq <= 257; seq += 1) {
      const noop = await original.call.edit(command(seq, 'setReadOnly', { readOnly: false }));
      expect(noop.outcomes).toMatchObject([{ seq, status: 'noop' }]);
    }
    const checkpoint = await original.call.checkpoint();
    const recovery = decodeWorkbookRecovery(checkpoint.state);
    expect(recovery.statuses).toHaveLength(256);
    expect(recovery.refusedSeqs).toEqual([]);
    expect(recovery.refusedRanges).toEqual([[1, 1]]);
    await replacement.call.openRecovery(checkpoint);
    for (const [seq, afterSeq] of [[258, 1], [259, 258], [260, 2]] as const) {
      const input = { ...command(seq, 'editCell', { args: [0, 0, 0, '7'], target }), afterSeq };
      const uninterrupted = await original.call.edit(input);
      const restored = await replacement.call.edit(input);
      expect(restored.outcomes).toEqual(uninterrupted.outcomes);
      expect(restored.outcomes).toMatchObject([{ seq, status: seq === 260 ? 'applied' : 'cancelled',
        ...(seq === 260 ? {} : { reason: 'dependency-refused' }) }]);
    }
    expect(await replacement.call.save()).toEqual(await original.call.save());
    expect(decodeWorkbookRecovery((await replacement.call.checkpoint()).state).refusedRanges)
      .toEqual([[1, 1], [258, 259]]);
  } finally { await original.dispose(); await replacement.dispose(); }
});

test('direct applyEdits preserves admitted command frame barriers', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const original = session(true);
  const edited = session(true);
  try {
    const bytes = await workbookBytes();
    const request = { sheet: 0, viewport: { x: 0, y: 0, width: 160, height: 64 },
      generation: 9, zoom: 1, dpr: 1 };
    const results = [];
    for (const active of [original, edited]) {
      const opened = await active.call.open(bytes.slice(0));
      const context = opened.commandContext!;
      const target = { sheetId: opened.sheets[0]!.id, range: { kind: 'a1' as const, a1: 'A1' } };
      const framing = active.call.commandFrame(request);
      const direct = active === edited ? active.call.applyEdits({ expectVersion: context.basis.version,
        steps: [{ op: 'setCellInputs', target, inputs: [['5']] }] }) : undefined;
      const editing = active.call.edit({ session: context.session, incarnation: 1, seq: 1,
        basis: { epoch: context.basis.epoch, version: active === edited ? 1 : 0 }, expectVersion: true,
        kind: 'editCell', payload: { payload: { args: [0, 0, 0, '7'], target } } });
      const [frame, applied, appliedDirect] = await Promise.all([framing, editing, direct]);
      if (direct) expect(appliedDirect).toMatchObject({ ok: true, applied: true });
      expect(applied.outcomes).toMatchObject([{ seq: 1, status: 'applied' }]);
      expect(applied.version).toBe(active === edited ? 2 : 1);
      expect(frame).toMatchObject({ coversThroughSeq: 1, payload: { sequence: 1,
        version: applied.engineVersion, requestGeneration: 9, viewport: request.viewport } });
      if (!frame) throw new Error('Missing command frame');
      results.push({ coversThroughSeq: frame.coversThroughSeq, requestGeneration: frame.payload.requestGeneration,
        displayList: new TextDecoder().decode(frame.payload.displayList), outcomes: applied.outcomes });
      const checkpoint = await active.call.checkpoint();
      expect(checkpoint.version).toBe(applied.version);
      expect(decodeWorkbookRecovery(checkpoint.state).engineVersion).toBe(applied.engineVersion);
    }
    expect(results[1]).toEqual(results[0]);
    expect((await edited.call.cellInputs(0, 'A1')).cells[0]?.[0]?.input).toBe('7');
  } finally {
    try {
      await Promise.all([original.call.dispose(), edited.call.dispose()]);
    } finally {
      await original.dispose();
      await edited.dispose();
    }
  }
});

test('restored refused duplicates preserve outer outcomes and retained status isolation', async () => {
  const opened = handle();
  const identity = { session: 'book', incarnation: 'first' };
  const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const state = createWorkbookCommandState(identity, { epoch: 1, version: opened.version() }, target, ['sheet:0']);
  const options = { handle: opened, state, captureRecovery: () => new Uint8Array([1]),
    frame: () => { throw new Error('Unexpected frame'); } };
  const original = createWorkbookCommandHost(options);
  const command = (seq: number): WorkbookSharedCommand => ({ session: 'book', incarnation: 1, seq,
    basis: { epoch: 1, version: 0 }, kind: 'select',
    payload: { payload: { selection: { ...target, sheetId: 'missing' } } } });
  const commands = [command(1), { ...command(2), basis: { epoch: 1, version: -1 },
    expectVersion: true }, { ...command(3), afterSeq: 1 }];
  for (const input of commands) await original.handler(null, input);
  const captured = await original.checkpoint(null);
  const checkpoint = isTransferResult(captured) ? captured.value : captured;
  const recovery = decodeWorkbookRecovery(checkpoint.state);
  const retainedStatuses = structuredClone(recovery.statuses);
  const restored = restoreWorkbookCommandState(opened, recovery, { ...identity, incarnation: 'replacement' });
  const replacement = createWorkbookCommandHost({ ...options, state: restored,
    revision: recovery.revision, versions: recovery.versions, refusedSeqs: recovery.refusedSeqs,
    refusedRanges: recovery.refusedRanges, refusedOutcomes: recovery.refusedOutcomes });
  for (const input of commands) {
    const duplicate = { ...input, basis: { epoch: 1, version: 0 }, replay: true as const };
    const uninterruptedResult = await original.handler(null, duplicate);
    const uninterrupted = isTransferResult(uninterruptedResult) ? uninterruptedResult.value : uninterruptedResult;
    const result = await replacement.handler(null, duplicate);
    const ack = isTransferResult(result) ? result.value : result;
    expect(ack.outcomes).toEqual(uninterrupted.outcomes);
    expect(ack.outcomes[0]?.status).toBe(input.seq === 3 ? 'cancelled' : 'refused');
    const outcome = ack.outcomes[0];
    if (!outcome?.workbook) throw new Error('Missing retained refusal');
    outcome.reason = 'changed';
    outcome.workbook.reason = 'changed';
    outcome.affectedSeqs?.push(99);
    outcome.workbook.affectedSeqs?.push(99);
    const repeated = await replacement.handler(null, duplicate);
    expect((isTransferResult(repeated) ? repeated.value : repeated).outcomes).toEqual(uninterrupted.outcomes);
  }
  const recaptured = await replacement.checkpoint(null);
  const next = isTransferResult(recaptured) ? recaptured.value : recaptured;
  const saved = decodeWorkbookRecovery(next.state);
  expect(saved.statuses).toEqual(retainedStatuses);
  expect(saved.refusedOutcomes).toEqual(recovery.refusedOutcomes);
});

test('recovery validates compact refusal ranges', () => {
  const opened = handle();
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0']);
  state.throughSeq = 9000;
  const encode = (refusedRanges: [number, number][]) => encodeWorkbookRecovery(new Uint8Array([1]),
    state, 0, [[0, opened.version()]], [], undefined, undefined, { refusedRanges });
  const valid: [number, number][] = [[1, 257], [300, 9000]];
  expect(decodeWorkbookRecovery(encode(valid)).refusedRanges).toEqual(valid);
  const bounded = Array.from({ length: 4096 }, (_, index): [number, number] => [index * 2 + 1, index * 2 + 1]);
  expect(decodeWorkbookRecovery(encode(bounded)).refusedRanges).toHaveLength(4096);
  const invalid: [number, number][][] = [
    [[0, 1]], [[2, 1]], [[1, 9001]], [[3, 4], [1, 2]], [[1, 3], [3, 4]], [[1, 2], [3, 4]],
    [[1.5, 2]], [[1, Number.MAX_SAFE_INTEGER + 1]],
    Array.from({ length: 4097 }, (_, index): [number, number] => [index * 2 + 1, index * 2 + 1]),
  ];
  for (const ranges of invalid) expect(() => decodeWorkbookRecovery(encode(ranges))).toThrow('Invalid workbook recovery checkpoint');
});

test('direct edits validate pending command settlements before synchronization', async () => {
  const opened = handle();
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'first' },
    { epoch: 1, version: opened.version() }, selected, ['sheet:0']);
  const host = createWorkbookCommandHost({ handle: opened, state,
    frame: () => { throw new Error('Unexpected frame'); } });
  host.validateDirectEdits();
  const pending = host.handler(null, { session: 'book', incarnation: 1, seq: 1,
    basis: { epoch: 1, version: 0 }, kind: 'select', payload: { payload: { selection: selected } } });
  expect(() => host.validateDirectEdits()).toThrow('Direct edits require an unused command host');
  expect(state.throughSeq).toBe(0);
  await pending;
  expect(() => host.validateDirectEdits()).toThrow('Direct edits require an unused command host');
});

test('cascade composition accepts released engine layouts and codecs six and seven', () => {
  for (const codec of [3, 5, 6, 7]) {
    for (const collaborative of [true, false]) {
      expect(isCollaborativeWorkbookRecovery(engineHeader(JSON.stringify({ collaborative }), codec))).toBe(collaborative);
    }
  }
  for (const codec of [2, 4, 8]) {
    expect(() => isCollaborativeWorkbookRecovery(engineHeader(undefined, codec))).toThrow('Unsupported workbook recovery checkpoint');
  }
});
