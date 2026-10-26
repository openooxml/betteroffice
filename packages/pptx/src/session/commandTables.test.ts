import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as wasm from '../wasm/loader';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods } from './commandProtocol';

const identity = { session: 'tables', incarnation: 'tables-worker' };
const input = { inputTime: 1, undoGroup: 'table' };

async function deck(merged: boolean): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url)));
  const path = 'ppt/slides/slide1.xml';
  if (!merged) zip.file(path, (await zip.file(path)!.async('string')).replace(/ (gridSpan|rowSpan|hMerge|vMerge)="[^"]*"/gu, ''));
  return zip.generateAsync({ type: 'uint8array' });
}

async function open(options: { merged?: boolean; collaboration?: boolean } = {}) {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await wasm.initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  let handle: wasm.PresentationHandle | undefined;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => (handle = original(...args)));
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const rpc = createSessionClient<PresentationCommandMethods, PresentationCommandEvents>(pair.client,
    { methods: PRESENTATION_COMMAND_METHODS });
  const dispose = async () => { try { await rpc.dispose(); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } };
  try {
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    let state: PresentationCommandAck = await rpc.call.openCommandEditor(new Uint8Array(await deck(options.merged ?? false)).buffer, {
      ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }],
      ...(options.collaboration ? { collaboration: true, clientId: 8201 } : {}),
    });
    let seq = 0;
    return {
      table() {
        const slide = handle!.snapshot().slides[0]!;
        const shape = slide.shapes.find((shape) => (shape.graphic as { type?: string } | null)?.type === 'table')!;
        const graphic = shape.graphic as { rows: unknown[]; grid: number[] };
        return { slide: slide.id, shape: shape.id, rows: graphic.rows.length, columns: graphic.grid.length, grid: graphic.grid, width: shape.width };
      },
      canUndo: () => handle!.canUndo(),
      async run(intent: unknown) {
        seq += 1;
        state = await rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope);
        return state.outcomes[0];
      },
      dispose,
    };
  } catch (error) {
    await dispose();
    throw error;
  }
}

const invoke = (method: string, args: unknown[]) => ({ kind: 'invoke', payload: { ...input, method, args } });
const refused = { status: 'refused', reason: 'invalid-payload' };

test('table structure invokes apply, undo and redo in a solo session', async () => {
  const session = await open();
  try {
    const { slide, shape } = session.table();
    expect(await session.run(invoke('insertTableRow', [slide, shape, 0, 'after', ['New']]))).toMatchObject({ status: 'applied', committed: true });
    expect(await session.run(invoke('insertTableColumn', [slide, shape, 0, 'before']))).toMatchObject({ status: 'applied', committed: true });
    expect(await session.run(invoke('deleteTableRow', [slide, shape, 0]))).toMatchObject({ status: 'applied', committed: true });
    expect(await session.run(invoke('deleteTableColumn', [slide, shape, 0]))).toMatchObject({ status: 'applied', committed: true });
    expect(session.table()).toMatchObject({ rows: 3, columns: 3 });
    expect(await session.run({ kind: 'undo', payload: input })).toMatchObject({ status: 'applied', committed: true });
    expect(session.table()).toMatchObject({ rows: 3, columns: 4 });
    expect(await session.run({ kind: 'redo', payload: input })).toMatchObject({ status: 'applied', committed: true });
    expect(session.table()).toMatchObject({ rows: 3, columns: 3 });
  } finally { await session.dispose(); }
});

test('refused table targets are refused without faulting the host or writing history', async () => {
  const session = await open();
  try {
    const { slide, shape } = session.table();
    expect(await session.run(invoke('insertTableRow', [slide, shape, 0, 'after']))).toMatchObject({ status: 'applied', committed: true });
    for (const intent of [
      invoke('insertTableRow', [slide, shape, 99, 'after']),
      invoke('deleteTableColumn', [slide, shape, 99]),
      invoke('insertTableColumn', [slide, shape, 0, 'after', ['Two\nlines']]),
      invoke('insertTableColumn', [slide, shape, 0, 'after', [], 'wide']),
      invoke('deleteTableRow', [slide, 'missing', 0]),
    ]) expect(await session.run(intent)).toMatchObject(refused);
    expect(session.table()).toMatchObject({ rows: 4, columns: 3 });
    expect(await session.run({ kind: 'undo', payload: input })).toMatchObject({ status: 'applied', committed: true });
    expect(session.table()).toMatchObject({ rows: 3, columns: 3 });
    expect(session.canUndo()).toBe(false);
    for (let rows = 3; rows > 1; rows--) {
      expect(await session.run(invoke('deleteTableRow', [slide, shape, 0]))).toMatchObject({ status: 'applied', committed: true });
    }
    expect(await session.run(invoke('deleteTableRow', [slide, shape, 0]))).toMatchObject(refused);
    expect(await session.run(invoke('insertTableRow', [slide, shape, 0, 'after']))).toMatchObject({ status: 'applied', committed: true });
  } finally { await session.dispose(); }
});

test('merged tables and collaborative sessions refuse table structure and stay writable', async () => {
  for (const options of [{ merged: true }, { collaboration: true }]) {
    const session = await open(options);
    try {
      const { slide, shape } = session.table();
      for (const intent of [
        invoke('insertTableRow', [slide, shape, 0, 'after']),
        invoke('deleteTableRow', [slide, shape, 0]),
        invoke('insertTableColumn', [slide, shape, 0, 'before']),
        invoke('insertTableColumn', [slide, shape, 0, 'before', [], 'copy']),
        invoke('deleteTableColumn', [slide, shape, 0]),
      ]) expect(await session.run(intent)).toMatchObject(refused);
      expect(await session.run(invoke('insertSlide', [1]))).toMatchObject({ status: 'applied', committed: true });
    } finally { await session.dispose(); }
  }
});

test('copy column width commands preserve widths through undo and redo', async () => {
  const session = await open();
  try {
    const before = session.table();
    const { slide, shape, grid } = before;
    expect(await session.run(invoke('insertTableColumn', [slide, shape, 1, 'after', undefined, 'copy']))).toMatchObject({ status: 'applied', committed: true });
    const expected = [...grid.slice(0, 2), grid[1], ...grid.slice(2)];
    expect(session.table()).toMatchObject({ grid: expected, width: expected.reduce((sum, width) => sum + width, 0) });
    expect(await session.run({ kind: 'undo', payload: input })).toMatchObject({ status: 'applied', committed: true });
    expect(session.table()).toEqual(before);
    expect(await session.run({ kind: 'redo', payload: input })).toMatchObject({ status: 'applied', committed: true });
    expect(session.table().grid).toEqual(expected);
  } finally { await session.dispose(); }
});
