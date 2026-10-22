import { expect, test } from 'bun:test';
import { openXlsx } from '../src/xlsx';
import { implicitFontWorkbook } from './format-fixtures';
import type { XlsxDefinedNameOp } from '@betteroffice/xlsx';

test('defined names support typed ops, complete pagination, scopes and save/reopen', async () => {
  const book = await openXlsx(await implicitFontWorkbook());
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

test('defined name MCP registration supports synthetic workbooks and read-only discovery', async () => {
  const { registerXlsxTools } = await import('../src/xlsx-mcp');
  const book = await openXlsx(await implicitFontWorkbook());
  const tools = new Map<string, { run: (args: any) => any; readOnly: boolean }>();
  const register: Parameters<typeof registerXlsxTools>[0] = (name, _description, _shape, readOnly, run) => { tools.set(name, { run, readOnly }); };
  try {
    registerXlsxTools(register, { get: () => book });
    const edit = tools.get('xlsx_edit_defined_names')!;
    expect(edit.readOnly).toBe(false);
    const call = (name: string, args: object) => tools.get(name)!.run({ document: 'book', ...args }).structuredContent;
    const version = book.handle.version();
    call('xlsx_edit_defined_names', { version, edits: [{ action: 'set_defined_name', name: 'Rate', refersTo: '=8' }] });
    expect(call('xlsx_list_defined_names', {}).items).toEqual([{ name: 'Rate', refersTo: '8', scope: null, sheet: null, hidden: false }]);
    expect(() => call('xlsx_edit_defined_names', { version, edits: [{ action: 'delete_defined_name', name: 'Rate' }] })).toThrow('Workbook changed');
    call('xlsx_edit_defined_names', { version: book.handle.version(), edits: [{ action: 'set_defined_name', name: 'Rate', refersTo: '9' }] });
    expect(call('xlsx_list_defined_names', {}).items[0].refersTo).toBe('9');
    call('xlsx_edit_defined_names', { version: book.handle.version(), edits: [{ action: 'delete_defined_name', name: 'rate' }] });
    expect(call('xlsx_list_defined_names', {}).total).toBe(0);
    tools.clear();
    registerXlsxTools(register, { get: () => book }, true);
    expect(tools.has('xlsx_edit_defined_names')).toBe(false);
    expect(tools.get('xlsx_list_defined_names')!.readOnly).toBe(true);
  } finally { book.close(); }
});
