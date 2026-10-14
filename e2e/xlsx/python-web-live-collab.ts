import { expect } from 'bun:test';
import { assertWorkbookMirror } from '../assertions';

import { SHEET } from './context';
import type { XlsxScenario } from './context';
import { PythonWorker, fromBase64, pythonMissing, toBase64 } from '../python';

const ROW = 390;

const OPEN = `
def run(state, input, timed):
    with timed('import'):
        import betteroffice_xlsx as bx
    with timed('open'):
        book = bx.Workbook.open_collaborative(base64.b64decode(input['bytes']), client_id=input['clientId'])
    state['book'] = book
    with timed('stateVector'):
        vector = book.state_vector()
    return {'clientId': book.client_id, 'stateVector': base64.b64encode(vector).decode()}
`;

const APPLY = `
def run(state, input, timed):
    book = state['book']
    with timed('applyUpdate'):
        book.apply_update(base64.b64decode(input['update']))
    with timed('read'):
        values = {address: book.value(0, address) for address in input['read']}
        formulas = {address: book.formula(0, address) for address in input['read']}
    with timed('stateVector'):
        vector = book.state_vector()
    return {'values': values, 'formulas': formulas, 'stateVector': base64.b64encode(vector).decode()}
`;

const EDIT = `
def run(state, input, timed):
    book = state['book']
    with timed('set'):
        for address, value in input['edits']:
            book.set(0, address, value)
    with timed('diff'):
        update = book.diff(base64.b64decode(input['peerVector']))
    with timed('stateVector'):
        vector = book.state_vector()
    return {'update': base64.b64encode(update).decode(), 'stateVector': base64.b64encode(vector).decode()}
`;

export const pythonWebLiveCollab: XlsxScenario = {
  name: 'python-web-live-collab',
  description:
    'A web replica and a Python replica of the same workbook exchange Yrs updates in both directions, including a concurrent edit to one cell, until both SDKs report the same inputs and values.',
  participants: ['web', 'python'],
  requires: pythonMissing,
  async run(ctx) {
    const { recorder, bytes, open } = ctx;
    const web = recorder.as('web');
    const handle = web.load(() =>
      open(bytes, { collaborative: true, clientId: 1 })
    );
    const a1 = (row: number, col: number) => handle.cell(SHEET, row, col).a1;
    const cells = [a1(ROW, 0), a1(ROW, 1), a1(ROW + 1, 0)];

    const worker = new PythonWorker(recorder);
    try {
      await worker.start();
      const opened = await worker.call<{
        clientId: number;
        stateVector: string;
      }>(
        'python:openCollaborative',
        OPEN,
        { bytes: toBase64(bytes), clientId: 2 },
        { bytes: bytes.byteLength }
      );
      expect(opened.clientId).toBe(2);

      expect(
        web.op('editCell:web', () =>
          handle.editCell(SHEET, ROW, 0, 'web says hi')
        ).applied
      ).toBe(true);
      expect(
        web.op('editCell:webFormula', () =>
          handle.editCell(SHEET, ROW + 1, 0, `=LEN(${a1(ROW, 0)})`)
        ).applied
      ).toBe(true);
      let update = web.op(
        'encodeStateAsUpdate',
        () => handle.encodeStateAsUpdate(fromBase64(opened.stateVector)),
        undefined,
        {
          bytes: handle.encodeStateAsUpdate(fromBase64(opened.stateVector))
            .byteLength,
        }
      );
      let mirrored = await worker.call<{
        values: Record<string, unknown>;
        formulas: Record<string, string | null>;
        stateVector: string;
      }>(
        'python:applyUpdate',
        APPLY,
        { update: toBase64(update), read: cells },
        { bytes: update.byteLength }
      );
      expect(mirrored.values[a1(ROW, 0)]).toBe('web says hi');
      expect(mirrored.values[a1(ROW + 1, 0)]).toBe(11);
      expect(mirrored.formulas[a1(ROW + 1, 0)]).toBe(`LEN(${a1(ROW, 0)})`);

      const back = await worker.call<{ update: string; stateVector: string }>(
        'python:editAndDiff',
        EDIT,
        {
          edits: [[a1(ROW, 1), 'python says hi']],
          peerVector: toBase64(handle.encodeStateVector()),
        }
      );
      const fromPython = fromBase64(back.update);
      expect(
        web.op(
          'applyUpdate:fromPython',
          () => handle.applyUpdate(fromPython),
          undefined,
          { bytes: fromPython.byteLength }
        ).applied
      ).toBe(true);
      expect(handle.cell(SHEET, ROW, 1).input).toBe('python says hi');

      expect(
        web.op('editCell:concurrent', () =>
          handle.editCell(SHEET, ROW + 2, 0, 'web wins?')
        ).applied
      ).toBe(true);
      const concurrent = await worker.call<{
        update: string;
        stateVector: string;
      }>('python:concurrentEdit', EDIT, {
        edits: [[a1(ROW + 2, 0), 'python wins?']],
        peerVector: toBase64(handle.encodeStateVector()),
      });
      const contested = fromBase64(concurrent.update);
      expect(
        web.op(
          'applyUpdate:concurrent',
          () => handle.applyUpdate(contested),
          undefined,
          { bytes: contested.byteLength }
        ).applied
      ).toBe(true);
      update = web.op('encodeStateAsUpdate:catchUp', () =>
        handle.encodeStateAsUpdate(fromBase64(concurrent.stateVector))
      );
      mirrored = await worker.call<{
        values: Record<string, unknown>;
        formulas: Record<string, string | null>;
        stateVector: string;
      }>(
        'python:applyUpdate',
        APPLY,
        { update: toBase64(update), read: [...cells, a1(ROW + 2, 0)] },
        { bytes: update.byteLength }
      );

      const resolved = handle.cell(SHEET, ROW + 2, 0).input;
      expect(['web wins?', 'python wins?']).toContain(resolved);
      expect(mirrored.values[a1(ROW + 2, 0)]).toBe(resolved);
      assertWorkbookMirror(handle, mirrored, [
        { row: ROW, col: 0, value: 'web says hi' },
        { row: ROW, col: 1, value: 'python says hi' },
        { row: ROW + 1, col: 0, value: 11 },
        { row: ROW + 2, col: 0, value: resolved },
      ]);
    } finally {
      await worker.close();
    }
  },
};
