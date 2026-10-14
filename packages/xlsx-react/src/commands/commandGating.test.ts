import { describe, expect, test } from 'bun:test';
import { XLSX_COMMAND_DESCRIPTORS, XLSX_COMMAND_IDS } from './descriptors';
import { evaluateXlsxCommand, type XlsxCommandEnvironment } from './evaluate';
import { PLAIN_FORMATTING, testEnvironment } from './testing';
import type { XlsxCommandArgs, XlsxCommandId, XlsxCommandState } from './types';

function code(state: XlsxCommandState): string | null {
  return state.enabled ? null : state.disabledReason.code;
}

function evaluate<K extends XlsxCommandId>(
  id: K,
  args?: XlsxCommandArgs[K],
  overrides: Partial<XlsxCommandEnvironment> = {}
) {
  return evaluateXlsxCommand(id, args, testEnvironment(overrides));
}

function range(top: number, left: number, bottom: number, right: number, merged = 0) {
  return {
    sheet: 0,
    top,
    left,
    bottom,
    right,
    merged: Array.from({ length: merged }, () => ({
      start: { row: top, col: left },
      end: { row: top, col: left + 1 },
    })),
    formatting: PLAIN_FORMATTING,
  };
}

const CELL_COMMANDS: readonly XlsxCommandId[] = [
  'bold',
  'italic',
  'strikethrough',
  'paintFormat',
  'fontFamily',
  'fontSize',
  'fontSizeStep',
  'textColor',
  'fillColor',
  'numberFormat',
  'decimalPlaces',
  'borderPreset',
  'borderStyle',
  'borderColor',
  'merge',
  'horizontalAlignment',
  'verticalAlignment',
  'textWrapping',
];

describe('xlsx command gating', () => {
  test('every command is available in a writable workbook with cells selected', () => {
    for (const id of XLSX_COMMAND_IDS) {
      if (id === 'searchMenus') continue;
      expect([id, code(evaluate(id))]).toEqual([id, null]);
    }
  });

  test('document writes stop in read-only mode; viewing commands stay available', () => {
    for (const id of XLSX_COMMAND_IDS) {
      const state = evaluate(id, undefined, { readOnly: true });
      const writes = XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument || id === 'proposalReject';
      if (id === 'searchMenus') continue;
      expect([id, code(state)]).toEqual([id, writes ? 'read-only' : null]);
    }
    const bold = evaluate('bold', undefined, { readOnly: true });
    expect(bold.active).toBe(false);
  });

  test('lifecycle reasons come before everything else except zoom', () => {
    for (const [status, reason] of [
      ['loading', 'document-loading'],
      ['empty', 'no-document'],
    ] as const) {
      for (const id of XLSX_COMMAND_IDS) {
        if (id === 'searchMenus') continue;
        const expected = id === 'zoom' ? null : reason;
        expect([id, code(evaluate(id, undefined, { status, readOnly: true }))]).toEqual([
          id,
          expected,
        ]);
      }
    }
    expect(code(evaluateXlsxCommand('bold', undefined, null))).toBe('editor-unavailable');
  });

  test('cell commands need cells, not a chart', () => {
    for (const id of CELL_COMMANDS) {
      expect([id, code(evaluate(id, undefined, { selection: null }))]).toEqual([
        id,
        'cell-selection-required',
      ]);
      expect([id, code(evaluate(id, undefined, { selection: 'chart' }))]).toEqual([
        id,
        'unsupported-selection',
      ]);
    }
    expect(code(evaluate('undo', undefined, { selection: 'chart' }))).toBeNull();
    expect(code(evaluate('save', undefined, { selection: null }))).toBeNull();
  });

  test('merge variants follow the selection shape and collaboration', () => {
    const cases: [ReturnType<typeof range>, Record<string, string | null>][] = [
      [range(0, 0, 0, 0), {
        all: 'multiple-cells-required',
        horizontal: 'multiple-columns-required',
        vertical: 'multiple-rows-required',
        unmerge: 'no-merged-cells',
      }],
      [range(0, 0, 0, 3), { all: null, horizontal: null, vertical: 'multiple-rows-required', unmerge: 'no-merged-cells' }],
      [range(0, 0, 4, 0), { all: null, horizontal: 'multiple-columns-required', vertical: null, unmerge: 'no-merged-cells' }],
      [range(0, 0, 4, 3, 1), { all: null, horizontal: null, vertical: null, unmerge: null }],
      [range(2, 2, 2, 2, 1), { all: 'multiple-cells-required', horizontal: 'multiple-columns-required', vertical: 'multiple-rows-required', unmerge: null }],
    ];
    for (const [selection, expected] of cases) {
      for (const [value, reason] of Object.entries(expected)) {
        const state = evaluate('merge', { value } as XlsxCommandArgs['merge'], { selection });
        expect([selection, value, code(state)]).toEqual([selection, value, reason]);
      }
      const overall = evaluate('merge', undefined, { selection });
      expect(overall.value).toEqual({
        rows: selection.bottom - selection.top + 1,
        columns: selection.right - selection.left + 1,
      });
      expect(code(overall)).toBe(
        Object.values(expected).some((reason) => reason === null) ? null : 'multiple-cells-required'
      );
    }
    for (const value of ['all', 'horizontal', 'vertical', 'unmerge'] as const) {
      const state = evaluate('merge', { value }, { collaborative: true, selection: range(0, 0, 3, 3, 1) });
      expect(code(state)).toBe('collaboration-unsupported');
    }
    expect(code(evaluate('merge', undefined, { collaborative: true }))).toBe(
      'collaboration-unsupported'
    );
  });

  test('marks keep mixed and unread formatting apart', () => {
    const mixed = range(0, 0, 1, 1);
    mixed.formatting = { ...PLAIN_FORMATTING, bold: undefined, fontSize: undefined, numberFormat: undefined };
    expect(evaluate('bold', undefined, { selection: mixed }).active).toBe('mixed');
    expect(evaluate('fontSize', undefined, { selection: mixed }).value).toBeNull();
    expect(evaluate('numberFormat', undefined, { selection: mixed }).value).toEqual({
      kind: null,
      pattern: null,
    });
    const unread = { ...mixed, formatting: null };
    expect(evaluate('bold', undefined, { selection: unread }).active).toBeUndefined();
    expect(evaluate('numberFormat', undefined, { selection: unread }).value).toBeNull();
  });

  test('options carry the state their arguments evaluate to', () => {
    const merge = evaluate('merge', undefined, { selection: range(0, 0, 0, 3) });
    expect(
      merge.options?.map((option) => [
        option.label,
        option.state.enabled ? null : option.state.disabledReason.code,
      ])
    ).toEqual([
      ['toolbar.merge.all', null],
      ['toolbar.merge.horizontal', null],
      ['toolbar.merge.vertical', 'multiple-rows-required'],
      ['toolbar.merge.unmerge', 'no-merged-cells'],
    ]);
    for (const option of evaluate('horizontalAlignment').options ?? []) {
      expect(option.state).toEqual(evaluate('horizontalAlignment', option.args));
    }
    const readOnly = evaluate('numberFormat', undefined, { readOnly: true });
    expect(readOnly.options?.every((option) => option.state.enabled === false)).toBe(true);
  });

  test('options report which choice applies', () => {
    const selection = range(0, 0, 0, 0);
    selection.formatting = { ...PLAIN_FORMATTING, numberFormat: 'currency', numberFormatPattern: '$#,##0.00' };
    expect(evaluate('numberFormat', undefined, { selection }).value).toEqual({
      kind: 'currency',
      pattern: '$#,##0.00',
    });
    expect(evaluate('numberFormat', { value: 'currency' }, { selection }).active).toBe(true);
    expect(evaluate('numberFormat', { value: 'percent' }, { selection }).active).toBe(false);
    expect(evaluate('horizontalAlignment', { value: 'left' }).active).toBe(true);
    expect(evaluate('fontFamily', { family: 'calibri' }).active).toBe(true);
    expect(evaluate('fontSize').options?.length).toBe(14);
    expect(evaluate('zoom', { scale: 1 }).active).toBe(true);
    expect(evaluate('borderStyle', undefined, { borderStyle: 'dashed' }).value).toBe('dashed');
    expect(evaluate('paintFormat', undefined, { paintFormat: true }).active).toBe(true);
  });

  test('arguments are validated', () => {
    const invalid: [XlsxCommandId, unknown][] = [
      ['fontSize', { points: 401 }],
      ['fontSize', { points: Number.NaN }],
      ['fontFamily', { family: ' ' }],
      ['textColor', { color: 'red' }],
      ['numberFormat', { value: 'accounting' }],
      ['merge', { value: 'diagonal' }],
      ['zoom', { scale: 0.1 }],
      ['decimalPlaces', { direction: 'sideways' }],
      ['proposalAccept', { proposalId: 'p1', force: 'yes' }],
      ['bold', { on: true }],
    ];
    for (const [id, args] of invalid) {
      expect([id, code(evaluate(id, args as never))]).toEqual([id, 'invalid-arguments']);
    }
    expect(code(evaluate('zoom', { scale: 4 }))).toBeNull();
    expect(code(evaluate('fontSize', { points: 1 }))).toBeNull();
  });

  test('history, export, search and proposals state their reasons', () => {
    expect(code(evaluate('undo', undefined, { canUndo: false }))).toBe('nothing-to-undo');
    expect(code(evaluate('undo', undefined, { canUndo: false, pendingInput: true }))).toBeNull();
    expect(code(evaluate('redo', undefined, { canRedo: false }))).toBe('nothing-to-redo');
    expect(code(evaluate('exportPng', undefined, { pngExport: false }))).toBe('png-unavailable');
    expect(code(evaluate('searchMenus'))).toBe('unsupported-command');
    expect(code(evaluate('proposalsPanel', undefined, { proposals: null }))).toBe(
      'unsupported-command'
    );
    expect(evaluate('proposalsPanel').value).toBe(1);
    expect(code(evaluate('proposalAccept', undefined, { proposals: [] }))).toBe('no-proposals');
    expect(code(evaluate('proposalAccept', { proposalId: 'p2' }))).toBe('proposal-not-found');
    expect(code(evaluate('proposalReject', { proposalId: 'p1' }, { readOnly: true }))).toBe(
      'read-only'
    );
    expect(evaluate('proposalAccept').options).toEqual([
      { args: { proposalId: 'p1' }, label: 'Audit agent', state: { enabled: true } },
    ]);
  });

  test('host-disabled commands say so', () => {
    const state = evaluate('undo', undefined, { hostDisabled: new Set(['undo']) });
    expect(code(state)).toBe('host-disabled');
  });
});
