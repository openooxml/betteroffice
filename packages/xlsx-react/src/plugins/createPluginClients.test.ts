import { expect, mock, test } from 'bun:test';
import { selectionAt, type WorkbookHandle } from '@betteroffice/xlsx';
import { createPeerlessPluginClients } from './createPeerlessPluginClients';
import { createXlsxPluginHost } from './createXlsxPluginHost';
import type { XlsxPluginEditorAccess } from './createPluginClients';
import type { XlsxPluginDocumentAccess } from './createPluginDocumentAccess';
import { defineXlsxPlugin } from './defineXlsxPlugin';
import type { XlsxPluginContext, XlsxPluginReadClient } from './types';

test('workbook navigation preserves argument reads and navigator arity', async () => {
  const calls: string[] = [];
  const handle = { version: () => 'v1', onUpdate: () => () => {} } as unknown as WorkbookHandle;
  const selectCells = mock((..._args: Parameters<XlsxPluginEditorAccess['navigator']['selectCells']>) => null);
  const scrollToCell = mock((..._args: Parameters<XlsxPluginEditorAccess['navigator']['scrollToCell']>) => null);
  const access: XlsxPluginEditorAccess = {
    handle: () => handle,
    admit: async (_handle, operation) => ({ ok: true, value: operation() }),
    commands: () => null,
    readOnlyRefusal: () => null,
    applyEdits: () => { throw new Error('Unavailable'); },
    navigator: { selectCells, scrollToCell },
  };
  let context!: XlsxPluginContext<null>;
  const host = createXlsxPluginHost({ ...access, geometry: () => null, translate: (key) => key });
  host.setPlugins([defineXlsxPlugin({ id: 'navigator', createState: () => null,
    initialize: (next) => { context = next; } })]);
  host.open(handle);
  const selection = { sheetId: 'sheet:0', selection: selectionAt({ row: 0, col: 0 }) };
  const selected = await context.navigation.selectCells(selection, {
    get expectVersion() { calls.push('version'); return 'v1'; },
    get focus() { calls.push('focus'); return true; },
  });
  expect(selected).toEqual({ ok: true });
  expect(calls).toEqual(['version', 'focus']);
  expect(selectCells.mock.calls[0]).toEqual([handle, selection, true, expect.any(Function)]);
  calls.length = 0;
  const cell = { sheetId: 'sheet:0', row: 0, col: 0 };
  expect(await context.navigation.scrollToCell(cell, {
    get expectVersion() { calls.push('version'); return 'v1'; },
    get align() { calls.push('align'); return 'nearest' as const; },
  })).toEqual({ ok: true });
  expect(calls).toEqual(['version', 'align']);
  expect(scrollToCell.mock.calls[0]).toEqual([handle, cell, 'nearest', expect.any(Function)]);
  host.close('unmounted');
});

test('a fulfilled worker read refuses a replaced document before returning', async () => {
  let identity: object = {};
  let context!: XlsxPluginContext<null>;
  let resolve!: (value: Awaited<ReturnType<XlsxPluginReadClient['readCells']>>) => void;
  let started!: () => void;
  const dispatched = new Promise<void>((done) => { started = done; });
  const reply = new Promise<Awaited<ReturnType<XlsxPluginReadClient['readCells']>>>((done) => { resolve = done; });
  const access: XlsxPluginDocumentAccess = {
    identity: () => identity, version: () => 'v1', currentVersion: () => 'v1',
    subscribe: () => () => {},
    admit: async (_identity, operation) => ({ ok: true, value: operation() }),
    readVersion: () => ({ ok: true, version: 'v1' }),
    readCells: () => {
      started();
      return reply.then((value) => {
        queueMicrotask(() => { identity = {}; host.close('document-replaced'); });
        return value;
      });
    },
    findText: () => { throw new Error('Unavailable'); },
    validateEdits: () => { throw new Error('Unavailable'); },
    readOnlyRefusal: () => null,
    applyEdits: () => { throw new Error('Unavailable'); },
    commands: () => null,
    navigator: { selectCells: () => null, scrollToCell: () => null },
  };
  const host = createXlsxPluginHost({ ...access, geometry: () => null, translate: (key) => key }, createPeerlessPluginClients);
  host.setPlugins([defineXlsxPlugin({ id: 'reader', createState: () => null,
    initialize: (next) => { context = next; } })]);
  host.open(identity);
  const reading = context.read.readCells({ ranges: [] });
  await dispatched;
  resolve({ ok: true, version: 'v1', ranges: [], sheets: [],
    calculation: { cycleCells: [], limitedCells: [], truncated: false } });
  expect(await reading).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
});
