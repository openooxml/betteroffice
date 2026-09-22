import { expect } from 'bun:test';

import { SHEET, displayedText, profiledDisplayList, profiledEdit, profiledOps } from './context';
import type { XlsxScenario } from './context';

/** Well below every pinned sample's content, so the edits never collide with it. */
const ROW = 200;

export const editingSession: XlsxScenario = {
  name: 'editing-session',
  description: 'One editor types values and a formula, triggers a dependent recalc, inserts and deletes tracks, walks undo/redo, saves and reopens.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const info = recorder.op('sheetInfo', () => handle.sheetInfo());
    expect(info.sheetNames.length).toBeGreaterThan(0);

    const initial = profiledDisplayList(handle, recorder, 'displayList:initial');
    expect(initial.commands.length).toBeGreaterThan(0);
    expect(initial.grid).toBeDefined();

    const a = handle.cell(SHEET, ROW, 0).a1;
    const b = handle.cell(SHEET, ROW, 1).a1;
    const c = handle.cell(SHEET, ROW, 2).a1;

    expect(profiledEdit(handle, recorder, 'editCell:number', ROW, 0, '40').applied).toBe(true);
    expect(handle.cell(SHEET, ROW, 0).input).toBe('40');
    expect(profiledEdit(handle, recorder, 'editCell:number', ROW, 1, '2').applied).toBe(true);

    const formula = `=${a}*${b}`;
    expect(profiledEdit(handle, recorder, 'editCell:formula', ROW, 2, formula).applied).toBe(true);
    expect(handle.cell(SHEET, ROW, 2)).toMatchObject({ input: formula, isFormula: true });
    expect(displayedText(handle, recorder, 'searchText:formulaResult', '80', ROW, 2)).toContain('80');

    const dependent = profiledEdit(handle, recorder, 'editCell:dependentRecalc', ROW, 0, '50');
    expect(dependent.applied).toBe(true);
    expect(dependent.changed).toContain(c);
    expect(displayedText(handle, recorder, 'searchText:recalculated', '100', ROW, 2)).toContain('100');

    const shifted = `=${handle.cell(SHEET, ROW + 1, 0).a1}*${handle.cell(SHEET, ROW + 1, 1).a1}`;
    expect(profiledOps(handle, recorder, 'applyOps:insertRows', [{ type: 'insertRows', sheet: SHEET, at: 0, count: 1 }]).applied).toBe(true);
    expect(handle.cell(SHEET, ROW + 1, 2).input).toBe(shifted);

    expect(profiledOps(handle, recorder, 'applyOps:deleteCols', [{ type: 'deleteCols', sheet: SHEET, at: 3, count: 1 }]).applied).toBe(true);
    expect(handle.cell(SHEET, ROW + 1, 2).input).toBe(shifted);

    expect(recorder.op('undo:deleteCols', () => handle.undo()).applied).toBe(true);
    expect(recorder.op('undo:insertRows', () => handle.undo()).applied).toBe(true);
    expect(handle.cell(SHEET, ROW, 2).input).toBe(formula);
    expect(recorder.op('redo:insertRows', () => handle.redo()).applied).toBe(true);
    expect(handle.cell(SHEET, ROW + 1, 2).input).toBe(shifted);

    const history = recorder.op('historyState', () => handle.historyState());
    expect(history.undoDepth).toBeGreaterThan(0);
    expect(history.redoDepth).toBe(1);

    const edited = profiledDisplayList(handle, recorder, 'displayList:afterEdits');
    expect(edited.commands.length).toBeGreaterThan(0);

    const saved = recorder.op('save', () => handle.save());
    expect(saved.byteLength).toBeGreaterThan(0);

    const reopened = recorder.op('reopen', () => open(saved));
    expect(reopened.cell(SHEET, ROW + 1, 2).input).toBe(shifted);
    expect(displayedText(reopened, recorder, 'searchText:afterReopen', '100', ROW + 1, 2)).toContain('100');
    expect(reopened.sheetInfo().sheetNames).toEqual(info.sheetNames);
  },
};
