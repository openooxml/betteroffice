import { describe, expect, spyOn, test } from 'bun:test';
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
