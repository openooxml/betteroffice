import { expect } from 'bun:test';

import { SHEET, displayedText, profiledEdit } from './context';
import type { XlsxScenario } from './context';
import { PythonWorker, fromBase64, pythonMissing, toBase64 } from '../python';

const ROW = 380;

const OPEN = `
def run(state, input, timed):
    with timed('import'):
        import betteroffice_xlsx as bx
    with timed('open'):
        book = bx.Workbook.open_recalculated(base64.b64decode(input['bytes']))
    state['book'] = book
    with timed('read'):
        return {
            'sheets': book.sheet_names,
            'label': book.value(0, input['label']),
            'formula': book.formula(0, input['formula']),
            'computed': book.value(0, input['formula']),
        }
`;

const EDIT = `
def run(state, input, timed):
    book = state['book']
    with timed('set'):
        for address, value in input['edits']:
            book.set(0, address, value)
    with timed('read'):
        eager = book.value(0, input['probe'])
    with timed('recalculate'):
        calculation = book.recalculate()
    with timed('reread'):
        computed = book.value(0, input['probe'])
    with timed('save'):
        data = book.save()
    return {'changed': calculation.changed, 'eager': eager, 'computed': computed, 'bytes': base64.b64encode(data).decode()}
`;

export const pythonRoundtrip: XlsxScenario = {
  name: 'python-roundtrip',
  description:
    'The web core edits and saves a workbook, the Python binding reopens those bytes, reads the formula and its value back, writes its own inputs, recalculates and saves, and the web core reopens the result.',
  participants: ['web', 'python'],
  requires: pythonMissing,
  async run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const label = handle.cell(SHEET, ROW, 0).a1;
    const left = handle.cell(SHEET, ROW + 1, 0).a1;
    const right = handle.cell(SHEET, ROW + 1, 1).a1;
    const total = handle.cell(SHEET, ROW + 1, 2).a1;

    expect(
      profiledEdit(handle, recorder, 'editCell:label', ROW, 0, 'handoff')
        .applied
    ).toBe(true);
    expect(
      recorder.op('editCells:operands', () =>
        handle.editCells(SHEET, [
          { row: ROW + 1, col: 0, input: '6' },
          { row: ROW + 1, col: 1, input: '7' },
        ])
      ).applied
    ).toBe(true);
    expect(
      profiledEdit(
        handle,
        recorder,
        'editCell:total',
        ROW + 1,
        2,
        `=${left}*${right}`
      ).applied
    ).toBe(true);
    expect(
      displayedText(handle, recorder, 'searchText:total', '42', ROW + 1, 2)
    ).toContain('42');
    const handed = recorder.op('save', () => handle.save());

    const worker = new PythonWorker(recorder);
    try {
      await worker.start();
      const opened = await worker.call<{
        sheets: string[];
        label: string;
        formula: string;
        computed: number;
      }>(
        'python:open',
        OPEN,
        { bytes: toBase64(handed), label, formula: total },
        { bytes: handed.byteLength }
      );
      expect(opened.sheets).toEqual(handle.sheetInfo().sheetNames);
      expect(opened.label).toBe('handoff');
      expect(opened.formula).toBe(`${left}*${right}`);
      expect(opened.computed).toBe(42);

      const edited = await worker.call<{
        changed: number;
        eager: number;
        computed: number;
        bytes: string;
      }>('python:editAndSave', EDIT, {
        edits: [
          [left, '9'],
          [right, '8'],
          [label, 'python wrote this'],
        ],
        probe: total,
      });
      expect(edited.eager).toBe(72);
      expect(edited.changed).toBe(0);
      expect(edited.computed).toBe(72);

      const returned = fromBase64(edited.bytes);
      const reopened = recorder.op('reopen:pythonBytes', () => open(returned));
      expect(reopened.cell(SHEET, ROW, 0).input).toBe('python wrote this');
      expect(reopened.cell(SHEET, ROW + 1, 2).input).toBe(`=${left}*${right}`);
      expect(
        displayedText(
          reopened,
          recorder,
          'searchText:pythonTotal',
          '72',
          ROW + 1,
          2
        )
      ).toContain('72');
    } finally {
      await worker.close();
    }
  },
};
