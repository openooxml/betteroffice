import { describe, expect, test } from 'bun:test';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { XLSX_COMMAND_DESCRIPTORS } from '../commands/descriptors';
import { evaluateXlsxCommand, type XlsxCommandEnvironment } from '../commands/evaluate';
import { testEnvironment } from '../commands/testing';
import type { XlsxCommandStore } from '../commands/types';
import { xlsxContextMenuEntries, xlsxContextMenuSections } from './menus';

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
});
