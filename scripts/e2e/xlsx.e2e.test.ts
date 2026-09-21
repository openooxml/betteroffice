/**
 * xlsx end-to-end: open pinned corpus workbooks, run an editing session against
 * each and check the results, timing every operation across the wasm boundary
 * together with the stage breakdown the core reports for it.
 */

import { afterAll, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook } from '../../packages/xlsx/src/wasm/loader';
import type {
  DisplayListProfile,
  EditProfile,
  WorkbookHandle,
} from '../../packages/xlsx/src/wasm/loader';
import { loadSample, samplesFor } from './corpus';
import { SampleRecorder, e2eEnabled, finishFormat } from './harness';
import type { SampleRun, StageProfile } from './harness';

const WASM = resolve(import.meta.dir, '../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const VIEWPORT = { x: 0, y: 0, width: 1280, height: 800 };
/** Well below every pinned sample's content, so the edits never collide with it. */
const ROW = 200;
const SHEET = 0;

const suite = e2eEnabled() ? describe : describe.skip;
const runs: SampleRun[] = [];

suite('xlsx end-to-end', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));
  afterAll(() => finishFormat('xlsx', runs));

  for (const sample of samplesFor('xlsx')) {
    it(sample.id, async () => {
      const bytes = await loadSample(sample);
      const started = performance.now();
      const handle = openWorkbook(bytes);
      const recorder = new SampleRecorder('xlsx', sample.id, performance.now() - started);
      try {
        scenario(handle, recorder);
      } finally {
        handle.dispose();
      }
      runs.push(recorder.finish());
    });
  }
});

function editStages(profile: EditProfile): StageProfile {
  return {
    validate: profile.validateMs,
    apply: profile.applyMs,
    recalc: profile.recalcMs,
    result: profile.resultMs,
  };
}

function displayStages(profile: DisplayListProfile): StageProfile {
  return { build: profile.buildMs, encode: profile.encodeMs };
}

function scenario(handle: WorkbookHandle, recorder: SampleRecorder): void {
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
  expect(
    profiledOps(handle, recorder, 'applyOps:insertRows', [
      { type: 'insertRows', sheet: SHEET, at: 0, count: 1 },
    ]).applied
  ).toBe(true);
  expect(handle.cell(SHEET, ROW + 1, 2).input).toBe(shifted);

  expect(
    profiledOps(handle, recorder, 'applyOps:deleteCols', [
      { type: 'deleteCols', sheet: SHEET, at: 3, count: 1 },
    ]).applied
  ).toBe(true);
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

  const reopened = recorder.op('reopen', () => openWorkbook(saved));
  try {
    expect(reopened.cell(SHEET, ROW + 1, 2).input).toBe(shifted);
    expect(displayedText(reopened, recorder, 'searchText:afterReopen', '100', ROW + 1, 2)).toContain('100');
    expect(reopened.sheetInfo().sheetNames).toEqual(info.sheetNames);
  } finally {
    reopened.dispose();
  }
}

function profiledEdit(
  handle: WorkbookHandle,
  recorder: SampleRecorder,
  op: string,
  row: number,
  col: number,
  input: string
) {
  let profile: EditProfile | undefined;
  return recorder.op(
    op,
    () => {
      const result = handle.editCellProfiled(SHEET, row, col, input);
      profile = result.profile;
      return result;
    },
    () => editStages(profile!)
  );
}

function profiledOps(handle: WorkbookHandle, recorder: SampleRecorder, op: string, ops: unknown[]) {
  let profile: EditProfile | undefined;
  return recorder.op(
    op,
    () => {
      const result = handle.applyOpsProfiled(ops);
      profile = result.profile;
      return result;
    },
    () => editStages(profile!)
  );
}

function profiledDisplayList(handle: WorkbookHandle, recorder: SampleRecorder, op: string) {
  let profile: DisplayListProfile | undefined;
  return recorder.op(
    op,
    () => {
      const result = handle.displayListProfiled(VIEWPORT);
      profile = result.profile;
      return result.displayList;
    },
    () => displayStages(profile!)
  );
}

/** The formatted text the sheet shows at `row`/`col`, found through text search. */
function displayedText(
  handle: WorkbookHandle,
  recorder: SampleRecorder,
  op: string,
  query: string,
  row: number,
  col: number
): string {
  const matches = recorder.op(op, () => handle.searchText(query));
  const match = matches.find((m) => m.sheet === SHEET && m.row === row && m.col === col);
  expect(match, `${query} at row ${row} col ${col}`).toBeDefined();
  return match!.text;
}
