import { describe, expect, test } from 'bun:test';
import { createT, en } from '@betteroffice/xlsx-i18n';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { XLSX_COMMAND_DESCRIPTORS } from '../commands/descriptors';
import { evaluateXlsxCommand, type XlsxCommandEnvironment } from '../commands/evaluate';
import { testEnvironment } from '../commands/testing';
import type { XlsxCommandStore } from '../commands/types';
import { xlsxContextMenuEntries, xlsxContextMenuSections, type XlsxContextFreezeTarget } from './menus';

function store(overrides: Partial<XlsxCommandEnvironment> = {}) {
  const calls: unknown[] = [];
  const env = testEnvironment(overrides);
  const fake = {
    getDescriptor: (id: keyof typeof XLSX_COMMAND_DESCRIPTORS) => XLSX_COMMAND_DESCRIPTORS[id],
    getState: (id: never, args?: never) => evaluateXlsxCommand(id, args, env),
    execute: async (id: string, args: unknown) => { calls.push([id, args]); return { ok: true, status: 'executed' }; },
  } as unknown as XlsxCommandStore;
  return { fake, calls };
}

const labels = (entries: readonly OverflowMenuEntry[]): unknown[] => entries.map((entry) =>
  entry.kind === 'separator' ? '-' : entry.kind === 'submenu' ? [entry.label, labels(entry.entries)] :
    entry.kind === 'item' ? entry.label : entry.id);

const t = createT(en);

function freezeAt(index: number, rows: number, cols: number) {
  const calls: Array<[number, number]> = [];
  const target: XlsxContextFreezeTarget = { index, rows, cols, apply: (r, c) => { calls.push([r, c]); } };
  return { target, calls };
}

function viewEntries(target: 'columnHeader' | 'rowHeader', freeze: XlsxContextFreezeTarget): OverflowMenuEntry[] {
  const { fake } = store();
  return xlsxContextMenuSections(target, fake, t, freeze).find((section) => section.id === 'view')?.entries ?? [];
}

function select(entries: readonly OverflowMenuEntry[], label: string) {
  const entry = entries.find((candidate) => candidate.kind === 'item' && candidate.label === label);
  if (entry?.kind !== 'item') throw new Error(`expected a ${label} item`);
  entry.onSelect();
}

describe('xlsx context menus', () => {
  test('the cell menu follows Excel order and shows the grid shortcuts', () => {
    const { fake } = store();
    const entries = xlsxContextMenuEntries('cell', fake, (key) => (key === 'keys.ctrl' ? 'Ctrl' : key));
    expect(labels(entries)).toEqual([
      'commands.cut', 'commands.copy', 'commands.paste', 'commands.pastePlainText', '-',
      ['commands.insert', ['commands.insertRowsAbove', 'commands.insertRowsBelow', 'commands.insertColumnsLeft', 'commands.insertColumnsRight']],
      ['commands.delete', ['commands.deleteRows', 'commands.deleteColumns']], 'commands.clearContents', '-',
      ['commands.numberFormat', ['toolbar.numberFormats.automatic', 'toolbar.numberFormats.plainText', 'toolbar.numberFormats.number',
        'toolbar.numberFormats.percent', 'toolbar.numberFormats.scientific', 'toolbar.numberFormats.currency',
        'toolbar.numberFormats.date', 'toolbar.numberFormats.time']],
    ]);
    const copy = entries.find((entry) => entry.kind === 'item' && entry.label === 'commands.copy');
    expect(copy).toMatchObject({ disabled: false, shortcut: expect.stringMatching(/^(Ctrl\+C|⌘C)$/) });
  });

  test('items disable with the command reason and run through execute', () => {
    const { fake, calls } = store({ readOnly: true });
    const entries = xlsxContextMenuEntries('cell', fake, (key) => key);
    const cut = entries.find((entry) => entry.kind === 'item' && entry.label === 'commands.cut');
    expect(cut).toMatchObject({ disabled: true, description: 'commands.reasons.readOnly' });
    const insert = entries.find((entry) => entry.kind === 'submenu' && entry.id === 'insert');
    expect(insert).toMatchObject({ disabled: true });
    const copy = entries.find((entry) => entry.kind === 'item' && entry.label === 'commands.copy');
    if (copy?.kind === 'item') copy.onSelect();
    expect(calls).toEqual([['copy', null]]);
  });

  test('the sheet tab menu has named sections later contributions can join', () => {
    const { fake, calls } = store({ workbook: { sheets: ['A', 'B'], active: 1, clipboardRead: true } });
    const sections = xlsxContextMenuSections('sheetTab', fake, (key) => key);
    expect(sections.map((section) => section.id)).toEqual(['sheet', 'arrange']);
    const [left, right] = sections[1].entries;
    expect(left).toMatchObject({ label: 'commands.moveSheetLeft', disabled: false });
    expect(right).toMatchObject({ label: 'commands.moveSheetRight', disabled: true, description: 'commands.reasons.sheetAtEdge' });
    if (left.kind === 'item') left.onSelect();
    expect(calls).toEqual([['moveSheet', { direction: 'left' }]]);
  });

  test('header menus carry the axis insert and delete entries', () => {
    for (const [target, entries] of [
      ['columnHeader', ['commands.insertColumnsLeft', 'commands.insertColumnsRight', 'commands.deleteColumns']],
      ['rowHeader', ['commands.insertRowsAbove', 'commands.insertRowsBelow', 'commands.deleteRows']],
    ] as const) {
      const { fake, calls } = store();
      const sections = xlsxContextMenuSections(target, fake, (key) => key);
      expect(sections.map((section) => section.id)).toEqual(['cells', 'delete']);
      expect(labels(sections.flatMap((section) => section.entries))).toEqual([...entries]);
      const first = sections[0].entries[0];
      if (first.kind !== 'item') throw new Error('expected a header menu item');
      first.onSelect();
      expect(calls).toEqual([[target === 'columnHeader' ? 'insertColumns' : 'insertRows',
        target === 'columnHeader' ? { position: 'left' } : { position: 'above' }]]);
    }
  });

  test('a column header freezes the columns to its left', () => {
    const { target, calls } = freezeAt(1, 0, 0);
    const entries = viewEntries('columnHeader', target);
    expect(labels(entries)).toEqual(['Freeze column A']);
    select(entries, 'Freeze column A');
    expect(calls).toEqual([[0, 1]]);
  });

  test('a column header keeps the frozen rows and offers unfreezing', () => {
    const { target, calls } = freezeAt(3, 2, 1);
    const entries = viewEntries('columnHeader', target);
    expect(labels(entries)).toEqual(['Freeze columns A–C', 'Unfreeze panes']);
    select(entries, 'Freeze columns A–C');
    select(entries, 'Unfreeze panes');
    expect(calls).toEqual([[2, 3], [0, 0]]);
  });

  test('a column header hides a freeze that would not change the split', () => {
    const { target } = freezeAt(3, 0, 3);
    expect(labels(viewEntries('columnHeader', target))).toEqual(['Unfreeze panes']);
  });

  test('the first column offers only unfreezing, and only while frozen', () => {
    const clear = freezeAt(0, 0, 0);
    const { fake } = store();
    expect(xlsxContextMenuSections('columnHeader', fake, t, clear.target).map((section) => section.id))
      .toEqual(['cells', 'delete']);
    const rowsOnly = freezeAt(0, 2, 0);
    expect(labels(viewEntries('columnHeader', rowsOnly.target))).toEqual(['Unfreeze panes']);
  });

  test('a row header freezes the rows above it', () => {
    const single = freezeAt(1, 0, 0);
    const entries = viewEntries('rowHeader', single.target);
    expect(labels(entries)).toEqual(['Freeze row 1']);
    select(entries, 'Freeze row 1');
    expect(single.calls).toEqual([[1, 0]]);
  });

  test('a row header keeps the frozen columns and offers unfreezing', () => {
    const { target, calls } = freezeAt(3, 1, 2);
    const entries = viewEntries('rowHeader', target);
    expect(labels(entries)).toEqual(['Freeze rows 1–3', 'Unfreeze panes']);
    select(entries, 'Freeze rows 1–3');
    select(entries, 'Unfreeze panes');
    expect(calls).toEqual([[3, 2], [0, 0]]);
  });

  test('a row header hides a freeze that would not change the split', () => {
    const { target } = freezeAt(3, 3, 0);
    expect(labels(viewEntries('rowHeader', target))).toEqual(['Unfreeze panes']);
    const clear = freezeAt(0, 0, 0);
    expect(labels(viewEntries('rowHeader', clear.target))).toEqual([]);
    const columnsOnly = freezeAt(0, 0, 2);
    expect(labels(viewEntries('rowHeader', columnsOnly.target))).toEqual(['Unfreeze panes']);
  });

  test('omitting the freeze target leaves the header menu as before', () => {
    const { fake } = store();
    for (const [target, expected] of [
      ['columnHeader', ['commands.insertColumnsLeft', 'commands.insertColumnsRight', 'commands.deleteColumns']],
      ['rowHeader', ['commands.insertRowsAbove', 'commands.insertRowsBelow', 'commands.deleteRows']],
    ] as const) {
      const sections = xlsxContextMenuSections(target, fake, (key) => key);
      expect(sections.map((section) => section.id)).toEqual(['cells', 'delete']);
      expect(labels(sections.flatMap((section) => section.entries))).toEqual([...expected]);
      expect(labels(xlsxContextMenuEntries(target, fake, (key) => key)))
        .toEqual([expected[0], expected[1], '-', expected[2]]);
    }
  });
});
