import { describe, expect, spyOn, test } from 'bun:test';
import { DEFAULT_MAX_FRAME_BYTES } from '../collaboration/protocol';
import { validateWorkbookCommandPayload, WorkbookCommandRefusal, type WorkbookCommand } from './commandProtocol';

function batch(steps: unknown): WorkbookCommand {
  return { session: 'book', incarnation: 'worker', seq: 1, basis: { epoch: 1, version: 'v1' },
    kind: 'applyEdits', payload: { args: [{ expectVersion: 'v1', steps }] } } as WorkbookCommand;
}

const target = { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } };
const step = (inputs: unknown) => ({ op: 'setCellInputs', target, inputs });

describe('bounded workbook batches', () => {
  for (const field of ['steps', 'rows', 'columns', 'guards', 'args'] as const) {
    for (const length of [3, 0xffff_ffff]) {
      test(`refuses sparse ${field} of length ${length} before materializing`, () => {
        const sparse = new Array(length);
        const command = field === 'steps' ? batch(sparse) :
          field === 'guards' ? batch([{ ...step([['value']]), expect: { cells: sparse } }]) :
          batch([step(field === 'rows' ? sparse : [sparse])]);
        if (field === 'args') command.payload = { args: sparse } as WorkbookCommand['payload'];
        let materialized = false;
        const original = Array.from;
        const converting = spyOn(Array, 'from').mockImplementation(((value: unknown, ...args: unknown[]) => {
          if (value === sparse) {
            materialized = true;
            throw new Error('Sparse array reached allocation');
          }
          return Reflect.apply(original, Array, [value, ...args]);
        }) as typeof Array.from);
        try {
          expect(() => validateWorkbookCommandPayload(command)).toThrow(WorkbookCommandRefusal);
          expect(materialized).toBe(false);
        } finally { converting.mockRestore(); }
      });
    }
  }

  test('refuses dense matrices exceeding the aggregate traversal budget', () => {
    const inputs = Array(1024).fill(Array(1024).fill('value'));
    expect(() => validateWorkbookCommandPayload(batch([step(inputs)]))).toThrow(WorkbookCommandRefusal);
  });

  test('refuses deeply nested extra fields', () => {
    let extra: unknown = null;
    for (let index = 0; index < 128; index++) extra = { extra };
    expect(() => validateWorkbookCommandPayload(batch([{ ...step([['value']]), extra }]))).toThrow(WorkbookCommandRefusal);
  });

  test('accepts a bounded dense batch', () => {
    expect(() => validateWorkbookCommandPayload(batch([step([['one', 'two']])]))).not.toThrow();
  });
});

describe('collaboration pipe payloads', () => {
  const command = (payload: unknown) => ({ ...batch([]), kind: 'collabPipe', payload }) as WorkbookCommand;
  const cursor = { sheet: 'Sheet', anchor: { row: 0, col: 0 }, head: { row: 1, col: 2 } };

  test('accepts XLSX cursors and provider options', () => {
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [
      { op: 'create', seq: 1, options: { maxFrameBytes: 1024, maxMessagesPerFrame: 16, user: { name: 'Ada' } } },
      { op: 'cursor', seq: 2, cursor, broadcast: true },
      { op: 'event', seq: 3, event: { type: 'message', data: new ArrayBuffer(4) } },
    ] }))).not.toThrow();
  });

  test('accepts configured frame limits above the default across pipe commands', () => {
    const maxFrameBytes = DEFAULT_MAX_FRAME_BYTES + 1024;
    const event = { op: 'event', seq: 2, event: { type: 'message', data: new ArrayBuffer(maxFrameBytes) } };
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [
      { op: 'create', seq: 1, options: { maxFrameBytes } }, event,
    ] }))).not.toThrow();
    const limits = new Map([[1, maxFrameBytes]]);
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [event] }), limits)).not.toThrow();
    expect(() => validateWorkbookCommandPayload(command({ provider: 2, ops: [event] }), limits)).toThrow(WorkbookCommandRefusal);
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [event] }))).toThrow(WorkbookCommandRefusal);
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [
      { op: 'destroy', seq: 1 }, event,
    ] }), limits)).toThrow(WorkbookCommandRefusal);
  });

  test('refuses frames above the configured limit and ignores replacement create limits', () => {
    const event = { op: 'event', seq: 2, event: { type: 'message', data: new ArrayBuffer(1025) } };
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [
      { op: 'create', seq: 1, options: { maxFrameBytes: 1024 } }, event,
    ] }))).toThrow(WorkbookCommandRefusal);
    expect(() => validateWorkbookCommandPayload(command({ provider: 1, ops: [
      { op: 'create', seq: 1, options: { maxFrameBytes: 2048 } }, event,
    ] }), new Map([[1, 1024]]))).toThrow(WorkbookCommandRefusal);
  });

  test('refuses invalid cursors and options that belong on main or to DOCX', () => {
    for (const options of [{ maxPendingBytes: 1024 }, { syncChunkBytes: 128 }, { maxSyncBytes: 1024 }]) {
      expect(() => validateWorkbookCommandPayload(command({ provider: 1,
        ops: [{ op: 'create', seq: 1, options }] }))).toThrow(WorkbookCommandRefusal);
    }
    for (const invalid of [{ ...cursor, sheet: '' }, { ...cursor, head: { row: -1, col: 0 } },
      { story: 'body', anchor: Uint8Array.of(0), head: Uint8Array.of(0) }]) {
      expect(() => validateWorkbookCommandPayload(command({ provider: 1,
        ops: [{ op: 'cursor', seq: 1, cursor: invalid, broadcast: true }] }))).toThrow(WorkbookCommandRefusal);
    }
  });
});
