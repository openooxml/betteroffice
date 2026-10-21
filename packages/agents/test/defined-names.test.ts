import { expect, test } from 'bun:test';
import { openXlsx } from '../src/xlsx';
import { xlsxFixture } from './format-fixtures';
import type { XlsxDefinedNameOp } from '@betteroffice/xlsx';

test('defined names support typed ops, complete pagination, scopes and save/reopen', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const ops: XlsxDefinedNameOp[] = [
      { type: 'setDefinedName', name: 'Rate', refers_to: '2', scope: null },
      { type: 'setDefinedName', name: 'Rate', refers_to: '3', scope: 0, hidden: true },
    ];
    book.handle.applyOps(ops);
    const page = book.handle.listDefinedNames({ limit: 1 });
    expect(page.items).toEqual([{ name: 'Rate', refersTo: '2', scope: null, hidden: false }]);
    expect(page.total).toBe(2);
    expect(page.nextOffset).toBe(1);
    expect(book.handle.listDefinedNames({ offset: 1 }).items[0]).toEqual({ name: 'Rate', refersTo: '3', scope: 0, hidden: true });
    const version = book.handle.version();
    book.editDefinedNames({ version, edits: [{ action: 'set_defined_name', name: 'rate', refersTo: '=4', sheet: 'sheet:0' }] });
    expect(book.listDefinedNames({ sheet: 'sheet:0' }).items[0]).toMatchObject({ refersTo: '4', hidden: true, sheet: 'sheet:0' });
    expect(() => book.editDefinedNames({ version, edits: [{ action: 'delete_defined_name', name: 'Rate' }] })).toThrow('Workbook changed');
    const before = book.handle.listDefinedNames();
    expect(() => book.editDefinedNames({ version: book.handle.version(), edits: [{ action: 'delete_defined_name', name: 'Rate' }, { action: 'set_defined_name', name: 'A1', refersTo: '1' }] })).toThrow();
    expect(book.handle.listDefinedNames()).toEqual(before);
    const reopened = await openXlsx(await book.export());
    try { expect(reopened.handle.listDefinedNames().items).toEqual(before.items); }
    finally { reopened.close(); }
    book.editDefinedNames({ version: book.handle.version(), edits: [{ action: 'delete_defined_name', name: 'RATE', sheet: 'sheet:0' }, { action: 'delete_defined_name', name: 'rate' }] });
    expect(book.listDefinedNames().total).toBe(0);
  } finally { book.close(); }
});
